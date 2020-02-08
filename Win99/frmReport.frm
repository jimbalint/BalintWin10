VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "1099 Report"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbForm 
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Default         =   -1  'True
      Height          =   615
      Left            =   2603
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   5843
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Form:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Year:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
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
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   9375
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim rs99 As New ADODB.Recordset
Dim rsBox As New ADODB.Recordset
Dim TaxYear As Integer

' 1=amt 2=tot / 0=tax 1-3 = column amt
Dim Amt(2, 3)

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    
    With Me
        
        .cmbForm.AddItem "ALL"
        .cmbForm.AddItem "1099-MISC"
        .cmbForm.AddItem "1099-R"
        .cmbForm.AddItem "1099-INT"
        .cmbForm.AddItem "1099-DIV"
        .cmbForm.ListIndex = 0
    
        PopTaxYear .cmbTaxYear
    
    End With
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub
Private Sub cmdPrint_Click()
    
    TaxYear = Me.cmbTaxYear.Text
    PrtInit ("Port")    ' "Port" = Portrait
    SetFont 9, Equate.Portrait
    
    Prvw.Caption = GLCompany.Name & " - 1099 Print " & TaxYear & " " & Form99.FormType

    With Me
        
        If .cmbForm.ListIndex = 0 Then
            PrintReport "1099-MISC"
            PrintReport "1099-R"
            PrintReport "1099-INT"
            PrintReport "1099-DIV"
        Else
            PrintReport .cmbForm.Text
        End If
    
    End With

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub PrintReport(ByVal FormTitle As String)

    Dim FormType As String
    Dim Ct99 As Integer
    FormType = Mid(FormTitle, 6)

    ' does any data exist?
    With Me
    
        SQLString = " SELECT * FROM Detail99 WHERE TaxYear = " & .cmbTaxYear.Text & _
                    " AND FormType = '" & FormType & "'"
        If Detail99.GetBySQL(SQLString) = False Then Exit Sub
                
        rs99.CursorLocation = adUseClient
        rs99.Fields.Append "PayeeID", adDouble
        rs99.Fields.Append "BoxName", adVarChar, 25, adFldIsNullable
        rs99.Fields.Append "FWT", adInteger
        rs99.Fields.Append "Amount", adCurrency
        rs99.Open , , adOpenDynamic, adLockOptimistic

        ' list of boxes in this run
        rsBox.CursorLocation = adUseClient
        rsBox.Fields.Append "BoxName", adVarChar, 100, adFldIsNullable
        rsBox.Fields.Append "BoxTitle", adVarChar, 100, adFldIsNullable
        rsBox.Open , , adOpenDynamic, adLockOptimistic

        Do
        
            ' it this an amount field?
            SQLString = " SELECT * FROM Field99 WHERE TaxYear = " & TaxYear & _
                        " AND FormType = '" & Detail99.FormType & "' " & _
                        " AND BoxName = '" & Detail99.BoxName & "'"
            If Field99.GetBySQL(SQLString) Then
                If Field99.FieldFormat = Equate.fmtString Then GoTo Next99
            Else
                GoTo Next99
            End If
        
            rs99.AddNew
            rs99!PayeeID = Detail99.PayeeID
            rs99!BoxName = Detail99.BoxName
            rs99!Amount = Detail99.FieldValue
            If InStr(1, LCase(Field99.FieldTitle), "tax withheld", vbTextCompare) Then
                rs99!FWT = 1
            Else
                rs99!FWT = 0
                rsBox.Find "BoxName = " & Trim(Detail99.BoxName), 0, adSearchForward, 1
                If rsBox.EOF Then
                
                    X = Detail99.BoxName
                    SQLString = " SELECT * FROM Field99 WHERE TaxYear = " & TaxYear & _
                                " AND FormType = '" & FormType & "' " & _
                                " AND BoxName = '" & Detail99.BoxName & "'"
                                
                    If Field99.GetBySQL(SQLString) Then
                        X = X & " " & Field99.FieldTitle
                    End If
                    X = Trim(X)
                
                    rsBox.AddNew
                    rsBox!BoxName = Detail99.BoxName
                    rsBox!BoxTitle = Detail99.BoxName & " " & Field99.FieldTitle
                    rsBox.Update
                End If
            End If
            rs99.Update
            
Next99:
            If Detail99.GetNext = False Then Exit Do
        
        Loop
                        
    End With

    ' ????
    ' rsBox.Sort = "BoxName"
    If rsBox.RecordCount > 3 Then
        MsgBox "This report not formatted for more than 3 non-tax boxes!", vbCritical
        GoBack
    End If
    
    ReportHeader FormTitle
    
    ' clear the totals
    Ct99 = 0
    For I = 0 To 3
        Amt(2, I) = 0
    Next I
    
    SQLString = " SELECT * FROM Payee99 ORDER BY PayeeName "
    If Payee99.GetBySQL(SQLString) = False Then GoBack      ' ????
    
    Do
        
        ' clear the amounts
        For I = 0 To 3
            Amt(1, I) = 0
        Next I
        
        ' amounts for this payee???
        rs99.Filter = "PayeeID = " & Payee99.PayeeID
        If rs99.RecordCount > 0 Then
            rs99.MoveFirst
            Do
                
                If rs99!FWT = 1 Then
                    Amt(1, 0) = rs99!Amount
                    Amt(2, 0) = Amt(2, 0) + rs99!Amount
                Else
                    If rsBox.RecordCount > 0 Then
                        I = 0
                        rsBox.MoveFirst
                        Do
                            I = I + 1
                            If rs99!BoxName = rsBox!BoxName Then
                                Amt(1, I) = rs99!Amount
                                Amt(2, I) = Amt(2, I) + rs99!Amount
                            End If
                            rsBox.MoveNext
                            If rsBox.EOF Then Exit Do
                        Loop
                    End If
                End If
                
                rs99.MoveNext
                If rs99.EOF Then Exit Do
            Loop
            
            ' print the line
            PrintValue(1) = Payee99.PayeeName:          FormatString(1) = "a25"
            PrintValue(2) = Payee99.PayeeNumber:        FormatString(2) = "n10"
            PrintValue(3) = " " & Payee99.FederalID:    FormatString(3) = "a16"
            
            ' tax amount
            PrintValue(4) = Amt(1, 0):                  FormatString(4) = "d12"
            
            ' other boxes
            I = 5
            If rsBox.RecordCount > 0 Then
                rsBox.MoveFirst
                Do
                    PrintValue(I) = Amt(1, I - 4)
                    FormatString(I) = "d14"
                    I = I + 1
                    rsBox.MoveNext
                    If rsBox.EOF Then Exit Do
                Loop
            End If
            
            PrintValue(I) = "":         FormatString(I) = "~"
            FormatPrint
            Ln = Ln + 1
            Ct99 = Ct99 + 1
            
            If Ln > MaxLines - 3 Then
                FormFeed
                ReportHeader FormTitle
            End If
            
        End If
        rs99.Filter = adFilterNone
                
        If Payee99.GetNext = False Then Exit Do
    
    Loop
    
    If Ln > MaxLines - 2 Then
        FormFeed
        ReportHeader FormTitle
    End If
    
    ' print totals
    Ln = Ln + 1
    PrintValue(1) = "TOTALS: ":             FormatString(1) = "a25"
    PrintValue(2) = "# PAYEES:":            FormatString(2) = "a10"
    PrintValue(3) = Ct99:                   FormatString(3) = "n16"
    PrintValue(4) = Amt(2, 0):              FormatString(4) = "d12"
    
    ' other boxes
    I = 5
    If rsBox.RecordCount > 0 Then
        rsBox.MoveFirst
        Do
            PrintValue(I) = Amt(2, I - 4)
            FormatString(I) = "d14"
            I = I + 1
            rsBox.MoveNext
            If rsBox.EOF Then Exit Do
        Loop
    End If
    
    PrintValue(I) = "":         FormatString(I) = "~"
    FormatPrint
    Ln = Ln + 1
    
    If Me.cmbForm.Text = "ALL" Then FormFeed

    rs99.Close
    rsBox.Close

End Sub

Private Sub ReportHeader(ByVal FormTitle As String)

    Columns = 105
    Y = "Tax Year: " & Me.cmbTaxYear.Text & " Form: " & FormTitle
    PageHeader GLCompany.Name, Y, "", "", 1
    Ln = Ln + 1
    
    PrintValue(1) = "NAME":                 FormatString(1) = "a25"
    PrintValue(2) = "NUMBER":               FormatString(2) = "r10"
    PrintValue(3) = " FED ID":              FormatString(3) = "a16"
    PrintValue(4) = "FED TAX":              FormatString(4) = "r12"
    
    K = 0
    J = 4
    
    If rsBox.RecordCount > 0 Then
        rsBox.MoveFirst
        Do
            J = J + 1
            PrintValue(J) = " " & Mid(rsBox!BoxTitle, 1, 11)
            FormatString(J) = "r14"
            rsBox.MoveNext
            If rsBox.EOF Then Exit Do
            K = K + 1
            If K = 4 Then Exit Do   ' only room for 3 fields besides fwt
        Loop
    End If
    
    J = J + 1
    PrintValue(J) = ""
    FormatString(J) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(Columns, "-"):                       FormatString(1) = "a" & Columns
    PrintValue(2) = " ":                                        FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1


End Sub


