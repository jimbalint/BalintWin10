VERSION 5.00
Begin VB.Form frmProgress 
   Caption         =   "Progress"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMsg3 
      Caption         =   "Message 3"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   8535
   End
   Begin VB.Label lblMsg2 
      Caption         =   "Message 2"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   8535
   End
   Begin VB.Label lblMsg1 
      Caption         =   "Message 1"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rct As Long
Dim SQLQuery As String
Dim BByte As Byte
Dim c As Currency
Dim i As Integer
Dim j As Integer
Dim w As String
Dim x As String
Dim y As Long
Dim z As Long
Dim dType As String
Dim AsciiChannel As Integer
Dim Ct As Long

Private Sub Form_Load()
   
   Set GLAccount = New cGLAccount
   Set GLAmount = New cGLAmount
   Set GLBranch = New cGLBranch
   Set GLCompany = New cGLCompany
   Set GLColumn = New cGLColumn
   Set GLDescription = New cGLDescription
   Set GLHistory = New cGLHistory
   Set GLJournal = New cGLJournal
   
   AsciiChannel = FreeFile
   Open "\balint\data\glx10001.txt" For Input As AsciiChannel
   
   CNOpen      ' open the gl .mdb file
   CNDesOpen   ' open the description file
   
   If Not GLAccount.GetData(10) Then
      MsgBox "Error: " & ErrMessage
   Else
      MsgBox "Find OK: " & GLAccount.Account
   End If
   
   c = GLAmount.GetAmount(GLAccount.Account, 2003, 1, 2)
   MsgBox ErrMessage & " " & c

   End

   
   ' clear the tables
   GLAccount.DeleteAll
   GLAmount.DeleteAll
   GLBranch.DeleteAll
   GLColumn.DeleteAll
   GLCompany.DeleteAll
   GLDescription.DeleteAll
   GLHistory.DeleteAll
   GLJournal.DeleteAll
   
   Do
      
      Input #AsciiChannel, dType
      
      Ct = Ct + 1
      If Ct Mod 20 = 0 Then
         lblMsg1.Caption = Count
         lblMsg2.Caption = dType
         frmProgress.Show
      End If
      
      Select Case dType
         Case "END"           ' End Of File
              MsgBox "All Done ..."
              End
         Case "CMP"           ' Company Info
              Company
         Case "ACT"           ' Account info
              Account
         Case "AMT"           ' Amount Info
              Amount
         Case "COL"           ' Column Info
              Column
         Case "BRA"           ' Branch Info
              Branch
         Case "JRN"           ' Journal Info
              Journal
         Case "DES"           ' Description
              Description
         Case "HIS"           ' History
              History
      End Select
   
   Loop
   
   MsgBox "All Done ..."

End Sub

Private Sub Company()
   
   GLCompany.Clear
   
   For i = 1 To 11
       
       Input #AsciiChannel, x
       
       If x <> "" Then
   
          If i = 1 Then GLCompany.Name = x
          If i = 2 Then GLCompany.LastUpdate = CLng(x)
          If i = 3 Then GLCompany.LastClose = CLng(x)
          If i = 4 Then GLCompany.RetEarnAcct = CLng(x)
          If i = 5 Then GLCompany.SuspAcct = CLng(x)
          If i = 6 Then GLCompany.NetProfitAcct = CLng(x)
          If i = 7 Then GLCompany.FirstPAcct = CLng(x)
          If i = 8 Then GLCompany.PctBaseAcct = CLng(x)
          If i = 9 Then GLCompany.SubDigits = CByte(x)
          If i = 10 Then GLCompany.NumberPds = CByte(x)
          If i = 11 Then GLCompany.FirstPd = CByte(x)
       
       End If
       
   Next i

   GLCompany.Save rAdd
               
End Sub

Private Sub Account()
   
   GLAccount.Clear
   
   For i = 1 To 10
      
      Input #AsciiChannel, x
      
      If x <> "" Then
         
         If i = 1 Then GLAccount.Account = CLng(x)
         If i = 2 Then GLAccount.AcctType = x
         
         ' description
         If i = 3 Then
            
            If Mid(x, 1, 1) = "," Then
               
               w = ""
               
               For y = 2 To Len(x)
                   If Mid(x, y, 1) = " " Then Exit For
                   w = w & Mid(x, y, 1)
               Next y
               
               GLAccount.DescNumber = CLng(w)
               GLAccount.Description = Mid(x, y + 1)
            
'        MsgBox x & " " & GLAccount.DescNumber & " " & GLAccount.Description
            
            Else
               
               GLAccount.DescNumber = 0
               GLAccount.Description = x
            
            End If
         
         End If
         
         If i = 4 Then GLAccount.TotalLevel = CByte(x)
         If i = 5 Then GLAccount.PrintTab = CByte(x)
         If i = 6 Then GLAccount.LineFeeds = CByte(x)
         If i = 7 Then GLAccount.BSColumn = CByte(x)
         If i = 8 Then
            BByte = CByte(x)
            If BByte And 2 ^ 7 Then GLAccount.AllStatements = True
            If BByte And 2 ^ 6 Then GLAccount.AllSchedules = True
            If BByte And 2 ^ 5 Then GLAccount.BranchAcct = True
            If BByte And 2 ^ 4 Then GLAccount.ConsAcct = True
            If BByte And 2 ^ 3 Then GLAccount.TotalOnLedger = True
            If BByte And 2 ^ 2 Then GLAccount.DollarSign = True
            If BByte And 2 ^ 1 Then GLAccount.SignRevStmt = True
            If BByte And 2 ^ 0 Then GLAccount.SignRevSched = True
         End If
'             If i = 9 Then glaccount.Date1 = CLng(x)
'             If i = 10 Then glaccount.Date2 = CLng(x)
             
      End If
      
   Next i
   
   GLAccount.Save rAdd
   
End Sub
Private Sub Amount()
   
   GLAmount.Clear
   
   For i = 1 To 15
      
      Input #AsciiChannel, x
      
      If x <> "" Then

         If i = 1 Then GLAmount.Account = CLng(x)
         If i = 2 Then GLAmount.FiscalYear = CLng(x)
         If i = 3 Then GLAmount.Amount01 = CDec(x)
         If i = 4 Then GLAmount.Amount02 = CDec(x)
         If i = 5 Then GLAmount.Amount03 = CDec(x)
         If i = 6 Then GLAmount.Amount04 = CDec(x)
         If i = 7 Then GLAmount.Amount05 = CDec(x)
         If i = 8 Then GLAmount.Amount06 = CDec(x)
         If i = 9 Then GLAmount.Amount07 = CDec(x)
         If i = 10 Then GLAmount.Amount08 = CDec(x)
         If i = 11 Then GLAmount.Amount09 = CDec(x)
         If i = 12 Then GLAmount.Amount10 = CDec(x)
         If i = 13 Then GLAmount.Amount11 = CDec(x)
         If i = 14 Then GLAmount.Amount12 = CDec(x)
         If i = 15 Then GLAmount.Amount13 = CDec(x)
         
      End If
   
   Next i
   
   GLAmount.Save rAdd
   
End Sub
Private Sub Column()

   GLColumn.Clear
   
   For i = 1 To 7
   
      Input #AsciiChannel, x
      
      If x <> "" Then

         If i = 1 Then GLColumn.ReportID = x
         If i = 2 Then GLColumn.ColumnNum = CByte(x)
         If i = 3 Then GLColumn.Description = x
         If i = 4 Then GLColumn.Value1 = CByte(x)
         If i = 5 Then GLColumn.Value2 = CByte(x)
         If i = 6 Then GLColumn.PrintTab = CByte(x)
         
         If i = 7 Then
            BByte = CByte(x)
            If BByte And 2 ^ 4 Then GLColumn.Column = True
            If BByte And 2 ^ 3 Then GLColumn.Percent = True
            If BByte And 2 ^ 2 Then GLColumn.NonPrint = True
            If BByte And 2 ^ 1 Then GLColumn.PriorYear = True
            If BByte And 2 ^ 0 Then GLColumn.Budget = True
         End If
             
      End If
   
   Next i
   
   GLColumn.Save rAdd

End Sub
Private Sub Branch()

End Sub
Private Sub Journal()
   
   GLJournal.Clear
   
   For i = 1 To 2
   
      Input #AsciiChannel, x
      
      If x <> "" Then
         If i = 1 Then GLJournal.JournalSource = CInt(x)
         If i = 2 Then GLJournal.JournalName = x
      End If
   Next i
   
   GLJournal.Save rAdd
   
End Sub
Private Sub Description()
   
   GLDescription.Clear
   rct = 0
   
   For i = 1 To 2
      Input #AsciiChannel, x
      If x <> "" Then
         If i = 1 Then GLDescription.Number = CLng(x)
         If i = 2 Then GLDescription.Description = x
      End If
   Next i
   
   GLDescription.Save rAdd
   
End Sub
Private Sub History()

   GLHistory.Clear
   
   For i = 1 To 10
      
      Input #AsciiChannel, x
   
      If x <> "" Then
         If i = 1 Then GLHistory.Account = CLng(x)
         If i = 2 Then GLHistory.FiscalYear = CLng(x)
         If i = 3 Then GLHistory.Period = CByte(x)
         If i = 4 Then GLHistory.Amount = CDec(x)
         If i = 5 Then GLHistory.Reference = x
         If i = 6 Then GLHistory.Description = x
         If i = 7 Then GLHistory.SourceCode = CByte(x)
         If i = 8 Then GLHistory.JournalSource = CByte(x)
         If i = 9 Then GLHistory.HisType = x
         If i = 10 Then
            BByte = CByte(x)
            If BByte And 2 ^ 0 Then GLHistory.UpdateFlag = True
         End If
      End If
   
   Next i
   
   GLHistory.Save rAdd
   
End Sub
