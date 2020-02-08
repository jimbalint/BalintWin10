VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
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
   ScaleHeight     =   9120
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   7080
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   4935
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   8655
      _cx             =   15266
      _cy             =   8705
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
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, k As Long
Dim X, Y, Z As String
Dim prt As Printer

Private Sub Command1_Click()
    PrintForm
End Sub

Private Sub Form_Load()

    For Each prt In Printers
        If InStr(1, Printer.DeviceName, "8400", vbTextCompare) Then
            Set Printer = prt
            Exit For
        End If
    Next
    
    Printer.NewPage
    Printer.CurrentX = 1000
    Printer.CurrentY = 1000
    Printer.Print "Hello ..."
    Printer.NewPage
    Printer.EndDoc
    End

'    End


'    PrtInit ("Port")
'    With Prvw.vsp
'        Debug.Print "Paper sizes available on the "; .Device; ":"
'        For I = 1 To 256
'          If .PaperSizes(I) Then Debug.Print " paper size "; I; " is available"
'        Next
'    End With
    
    
'    PrtInit ("Port")    ' "Port" = Portrait
'    SetFont 10, Equate.Portrait
'
'    With Prvw.vsp
'        .PaperSize = pprLegal
'        .PhysicalPage = False
'    End With
'
''    With Prvw.vsp
''        MsgBox .PageHeight
''        ' .PaperSize = pprLegal
''        MsgBox .PageHeight
''        .ShowGuides = gdShow
''        .PhysicalPage = False
''    End With
'
'    j = 14 * 6
'    For i = 1 To j
'        PrintValue(1) = "AAAAA":    FormatString(1) = "a10"
'        PrintValue(2) = i:          FormatString(2) = "n9"
'        PrintValue(3) = " ":        FormatString(3) = "~"
'        FormatPrint
'        Ln = Ln + 1
'    Next i
'
'    Prvw.vsp.EndDoc
'    Prvw.Show vbModal
    
  '  End

End Sub

'Sub Form_Click()
'   Dim CX, CY, Msg, XPos, YPos   ' Declare variables.
'   ScaleMode = 3   ' Set ScaleMode to
'         ' pixels.
'   DrawWidth = 5   ' Set DrawWidth.
'   ForeColor = QBColor(4)   ' Set foreground to red.
'   FontSize = 24   ' Set point size.
'   CX = ScaleWidth / 2   ' Get horizontal center.
'   CY = ScaleHeight / 2   ' Get vertical center.
'   Cls   ' Clear form.
'   Msg = "Happy New Year!"
'   CurrentX = CX - TextWidth(Msg) / 2   ' Horizontal position.
'   CurrentY = CY - TextHeight(Msg)   ' Vertical position.
'   Print Msg   ' Print message.
'   Do
'      XPos = Rnd * ScaleWidth   ' Get horizontal position.
'      YPos = Rnd * ScaleHeight   ' Get vertical position.
'      PSet (XPos, YPos), QBColor(Rnd * 15)   ' Draw confetti.
'      DoEvents   ' Yield to other
'   Loop   ' processing.
'End Sub

