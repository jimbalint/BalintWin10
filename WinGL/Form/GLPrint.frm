VERSION 5.00
Begin VB.Form frmGLPrint 
   Caption         =   "GL Print File Setup"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbEndPeriod 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cmbFiscalYear 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblEndPeriod 
      Caption         =   "End Period"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblFY 
      Caption         =   "Fiscal Year:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmGLPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndYMs(11) As Long
Private Sub cmbFiscalYear_Click()
    EndPeriodSet (CInt(cmbFiscalYear))
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
    
    ' !!!!!!! same for Trial Balance only !!!!!
    GLPrint.EndDate = EndYMs(cmbEndPeriod.ListIndex)
    GLPrint.BeginDate = GLPrint.EndDate
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Unload Me
End Sub

Private Sub Form_Load()
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim RetVal As Boolean
    
   ' assign the GLPrint values
   GLPrint.Clear

   GLPrint.RegBraCon = Equate.Regular     ' 1=regular  2=branch  3=consolidated  4=budget  AB%
   GLPrint.StaSch = Equate.Stmt           ' 1=statment  2=schedule
   GLPrint.RegCmp = Equate.Regular        ' 1=regular  2=comparative
   GLPrint.PrintBIB = Equate.PrtBoth      ' 1=bal sht  2=inc stmt  3=both

   GLPrint.LowerCaseDate = False
   GLPrint.PrtAcctNum = False
   GLPrint.PrtZeroBal = False
   GLPrint.RoundDollars = False
   GLPrint.SepPage = False
   GLPrint.SupprCP = False
   GLPrint.UseMathRec = True
   GLPrint.WidePrint = True

   GLPrint.LowAccount = 1
   GLPrint.HiAccount = 999999999

   GLPrint.LowBranchAcct = 1
   GLPrint.HiBranchAcct = 99

   GLPrint.LowConsAcct = 1
   GLPrint.HiConsAcct = 99       ' ?????

   GLPrint.Output = ""
   GLPrint.Copies = 1
   GLPrint.ReportDate = 20030630
   GLPrint.User = "JIM"

   ' '''''''''''''''
   GLPrint.BeginDate = 200306
   GLPrint.EndDate = 200306

'   ' >>>> get from glprint
'   FontName = "Arial"
'   FontSize = 12

   ' find how many fiscal years exists in glamount
   ' loop thru glamount for the first zero type glaccount
   GLAccount.GetFirst
   Do Until GLAccount.AcctType = "0"
      RetVal = GLAccount.GetNext(GLAccount.Account)
   Loop
   
   rs.Source = "select [FiscalYear] from GLAmount " & _
               "where Account = " & GLAccount.Account & _
               " order by FiscalYear desc"
                    
   Set rs.ActiveConnection = cn
        
   rs.Open
        
   If rs.EOF = True And rs.BOF = True Then
      MsgBox "No amount data ???"
      End
   End If

   Do Until rs.EOF = True
      cmbFiscalYear.AddItem rs.Fields("FiscalYear")
      rs.MoveNext
   Loop
   
   Set rs = Nothing
   
   ' default to the first entry
   cmbFiscalYear.ListIndex = 0

   EndPeriodSet (CInt(cmbFiscalYear))

End Sub

Private Sub EndPeriodSet(ByVal FY As Integer)
    
    Dim i As Integer
    Dim v As Variant
    
    cmbEndPeriod.Clear
      
    If GLCompany.FirstPd = 1 Then
       v = DateSerial(FY, GLCompany.FirstPd, 1)
    Else
       v = DateSerial(FY - 1, GLCompany.FirstPd, 1)
    End If

    cmbEndPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    EndYMs(0) = Year(v) * 100 + Month(v)
    
    For i = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbEndPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
        EndYMs(i) = Year(v) * 100 + Month(v)
    Next i
    
    cmbEndPeriod.ListIndex = 0
    
End Sub
