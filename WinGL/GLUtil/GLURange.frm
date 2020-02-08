VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmGLUrange 
   Caption         =   "GL Clear and Update"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GLURange.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLook 
      Height          =   375
      Left            =   9360
      Picture         =   "GLURange.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1560
      TabIndex        =   13
      Top             =   3360
      Width           =   7695
      Begin VB.OptionButton optReUpdate 
         Caption         =   "Clear and &RE-UPDATE History"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton optClear 
         Caption         =   "&Clear and &DELETE History"
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox cmbStartPeriod 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.ComboBox cmbEndPeriod 
      Height          =   360
      Left            =   6480
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ComboBox cmbFiscalYear 
      Height          =   360
      Left            =   5640
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin TDBNumber6Ctl.TDBNumber tdbSuspenseAcct 
      Height          =   375
      Left            =   5820
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   661
      Calculator      =   "GLURange.frx":0614
      Caption         =   "GLURange.frx":0634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "GLURange.frx":06A2
      Keys            =   "GLURange.frx":06C0
      Spin            =   "GLURange.frx":070A
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0 = Create"
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblSuspenseAcct 
      Alignment       =   2  'Center
      Caption         =   "Sus&pense Account:"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label lblStartPd 
      Caption         =   "&Start Period"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblEndPeriod 
      Caption         =   "&End Period:"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblFY 
      Caption         =   "&Fiscal Year:"
      Height          =   255
      Left            =   4140
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmGLUrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndYMs(11) As Long
Dim StartYMs(11) As Long
Dim jj As Integer
Dim ll As Long
Dim mr As Integer

Public bytFrame1 As Byte
Public bytFrame2 As Byte
Public bytFrame3 As Byte
Public bytFrame4 As Byte

' need to add to GLPrint !!!
Public FiscalYear As Long
Public StartPd As Integer
Public EndPD As Integer
Public JournalSource As Integer
Public FirstPeriod As Byte
Public BeginDate As Long
Public EndDate As Long

Dim xDB As New XArrayDB


Private Sub cmbFiscalYear_Click()
    EndPeriodSet (CInt(cmbFiscalYear))
End Sub

Private Sub cmdColumns_Click()
    frmGLColumn.Show vbModal
End Sub

Private Sub CmdExit_Click()
    GoBack
End Sub

Private Sub cmdLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbSuspenseAcct = frmAcctLookup.SelAcct
    Me.optReUpdate.SetFocus
End Sub

Private Sub cmdOK_Click()
    
Dim DeleteHist As Boolean
    
Dim FY As Integer
Dim Pd As Byte
    
    If Me.optClear Then
       mr = MsgBox("Are you SURE you want to DELETE the HISTORY ?", vbCritical + vbOKCancel + vbDefaultButton2, "DELETE HISTORY !!! ???")
       If mr = vbCancel Then
          Me.optClear = False
          Me.optReUpdate = True
          Exit Sub
       End If
    End If
    
    If Me.optClear Then
       DeleteHist = True
    Else
       DeleteHist = False
    End If
    
    StartPd = Me.cmbStartPeriod.ListIndex + 1
    EndPD = Me.cmbEndPeriod.ListIndex + 1
    
    x = "Fiscal Year: " & frmGLUrange.cmbFiscalYear
    
    FY = CInt(frmGLUrange.cmbFiscalYear)
    Pd = frmGLUrange.cmbStartPeriod.ListIndex + 1
    
    x = x & " Start " & PeriodName(FY, _
                                           Pd, _
                                           GLCompany.FirstPeriod, _
                                           GLCompany.NumberPds)
                                           
    Pd = frmGLUrange.cmbEndPeriod.ListIndex + 1
                                           
    x = x & " - End " & PeriodName(FY, _
                                   Pd, _
                                   GLCompany.FirstPeriod, _
                                   GLCompany.NumberPds)
                                           
                
    xDB.ReDim 0, 5, 0, 0
    xDB(1, 0) = GLCompany.Name & " Clear and Update"
    xDB(2, 0) = x
    xDB(3, 0) = " "
    xDB(4, 0) = String(60, "=")
    xDB(5, 0) = " "

    Set uDB = ClearGLAmount(frmGLUrange.cmbFiscalYear, _
                                    frmGLUrange.cmbFiscalYear, _
                                    frmGLUrange.StartPd, _
                                    frmGLUrange.EndPD, _
                                    DeleteHist)

    xDBAssign

    Set uDB = ClearGLBudget(frmGLUrange.cmbFiscalYear, _
                                    frmGLUrange.cmbFiscalYear, _
                                    frmGLUrange.StartPd, _
                                    frmGLUrange.EndPD)

    xDBAssign

    Set uDB = UpdateGLAmount(frmGLUrange.cmbFiscalYear, _
                             frmGLUrange.cmbFiscalYear, _
                             frmGLUrange.StartPd, _
                             frmGLUrange.EndPD, _
                             frmGLUrange.tdbSuspenseAcct, _
                             CompanyID)

    xDBAssign

    Set uDB = MathUpdate(frmGLUrange.cmbFiscalYear, _
                         frmGLUrange.cmbFiscalYear, _
                         frmGLUrange.StartPd, _
                         frmGLUrange.EndPD)
                                
    xDBAssign

            
TMP:
    Set frmResults = New frmResults
    frmResults.lblCompanyName = GLCompany.Name
    frmResults.lblMsg1 = "Clear and Update GL Amounts"
    frmResults.lblMsg2 = ""
    frmResults.lblMsg3 = ""
    For i = 1 To xDB.UpperBound(1)
        frmResults.List1.AddItem xDB(i, 0)
    Next i
    frmResults.Show vbModal
    
    GoBack

End Sub

Private Sub cmdOptions_Click()
   frmGLPrint2.Show vbModal
End Sub

Private Sub Form_Load()
   
Dim rs As ADODB.Recordset
Dim RetVal As Boolean

   Me.lblCompanyName = GLCompany.Name

   Set rs = New ADODB.Recordset
   rs.Source = "Select DISTINCT FiscalYear from GLAmount order by FiscalYear Desc"
   
   Set rs.ActiveConnection = cn
        
   rs.Open
        
   If rs.EOF = True And rs.BOF = True Then
      MsgBox "No amount data ???"
      End
   End If

   ll = 0
   jj = 0
   Do Until rs.EOF = True
      cmbFiscalYear.AddItem rs.Fields("FiscalYear")
      If rs!FiscalYear = GLPrint.FiscalYear Then jj = ll
      rs.MoveNext
      ll = ll + 1
   Loop
   cmbFiscalYear.ListIndex = jj
   
   Set rs = Nothing
   
   ' default to the first entry
   EndPeriodSet (CInt(cmbFiscalYear))

   If GLPrint.BeginDate = 0 Then
      Me.cmbStartPeriod.ListIndex = 0
   Else
      For jj = 0 To 11
         If StartYMs(jj) = GLPrint.BeginDate Then
            Me.cmbStartPeriod.ListIndex = jj
         End If
      Next jj
   End If
   
   If GLPrint.EndDate = 0 Then
      Me.cmbEndPeriod.ListIndex = 0
   Else
      For jj = 0 To 11
         If EndYMs(jj) = GLPrint.EndDate Then
            Me.cmbEndPeriod.ListIndex = jj
         End If
      Next jj
   End If
      
'   ' init the journal source combo
'   For ll = 1 To 10
'       If GLJournal.GetData(CLng(ll)) = True Then
'          cmbJournalSource.AddItem ll & " - " & GLJournal.JournalName
'       End If
'   Next ll
'   cmbJournalSource.ListIndex = 0
   
End Sub

Private Sub EndPeriodSet(ByVal FY As Integer)
    
    Dim i As Integer
    Dim v As Variant
    
    cmbEndPeriod.Clear
    cmbStartPeriod.Clear
      
    If GLCompany.FirstPeriod = 1 Then
       v = DateSerial(FY, GLCompany.FirstPeriod, 1)
    Else
       v = DateSerial(FY - 1, GLCompany.FirstPeriod, 1)
    End If

    cmbEndPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    cmbStartPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    EndYMs(0) = Year(v) * 100 + Month(v)
    StartYMs(0) = Year(v) * 100 + Month(v)
    
    For i = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbEndPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
        cmbStartPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
        EndYMs(i) = Year(v) * 100 + Month(v)
        StartYMs(i) = Year(v) * 100 + Month(v)
    Next i
    
    cmbEndPeriod.ListIndex = 0
    cmbStartPeriod.ListIndex = 0
    
End Sub

Private Sub txtLoAccount_GotFocus()
   txtLoAccount.SelStart = 0
   txtLoAccount.SelLength = Len(txtLoAccount.Text)
End Sub
Private Sub txtHiAccount_GotFocus()
   txtHiAccount.SelStart = 0
   txtHiAccount.SelLength = Len(txtHiAccount.Text)
End Sub
Private Sub txtLoBranch_GotFocus()
   txtLoBranch.SelStart = 0
   txtLoBranch.SelLength = Len(txtLoBranch.Text)
End Sub
Private Sub txtHiBranch_GotFocus()
   txtHiBranch.SelStart = 0
   txtHiBranch.SelLength = Len(txtHiBranch.Text)
End Sub
Private Sub txtLoCons_GotFocus()
   txtLoCons.SelStart = 0
   txtLoCons.SelLength = Len(txtLoCons.Text)
End Sub
Private Sub txtHiCOns_GotFocus()
   txtHiCons.SelStart = 0
   txtHiCons.SelLength = Len(txtHiCons.Text)
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

Private Sub xDBAssign()

Dim i As Integer
Dim j As Integer
    
    For i = 1 To uDB.UpperBound(1)
        xDB.AppendRows
        j = xDB.UpperBound(1)
        xDB(j, 0) = uDB(i, 0)
    Next i

End Sub

