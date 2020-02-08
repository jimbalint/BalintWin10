VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form frmPreview 
   Caption         =   "GL Statement Print"
   ClientHeight    =   3600
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   6825
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6825
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "PRINT ALL PAGES  (F10)"
      Default         =   -1  'True
      Height          =   855
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT (ESCAPE)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VSPrinter8LibCtl.VSPrinter vsp 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2325
      _cx             =   4101
      _cy             =   4048
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   7.36607142857143
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    If PrvwReturn = False Then
        GoBack
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
   
   Me.WindowState = vbMaximized
   Form_Resize
   
'   frmPreview.Height = Screen.Height
'   frmPreview.Width = frmPreview.vsp.Width + 1000
'   frmPreview.Top = 0
'   frmPreview.Left = 0
'
'   frmPreview.vsp.Height = Screen.Height - 500
'   frmPreview.vsp.Left = 500
'   frmPreview.vsp.Top = 0
'
'   frmPreview.vsp.EndDoc

    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
    Me.vsp.AbortWindow = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
        Case vbKeyF10: cmdPrintAll_Click
    End Select
    
End Sub

Private Sub Form_Resize()
    
    vsp.Height = Me.Height * 0.95
    vsp.Width = Me.Width * 0.9
    vsp.Left = 500
    vsp.Top = 0
    
    cmdExit.Left = Me.Width - 1300
    cmdExit.Top = 200

    With Me.cmdPrintAll
        .Left = Me.Width - 1300
        .Top = 1200
        .ToolTipText = "PRINT ALL PAGES TO DFLT PRINTER"
    End With

End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GoBack
End Sub

Private Sub cmdPrintAll_Click()
    Me.vsp.PrintDoc
End Sub


