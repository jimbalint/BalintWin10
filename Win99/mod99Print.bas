Attribute VB_Name = "mod99Print"
Option Explicit

Dim vPos, vPos1, hPos As Long
Dim I, J, K As Long
Dim X, Y, Z As String
Dim PayerDemo(5) As String

Public Sub TestPrint()

    PrtInit ("Port")    ' "Port" = Portrait
    SetFont 10, Equate.Portrait
    
    ' ******************************
    Nudge = 30
    
    ' ******************************

Dim Rw, Co As Integer

    For Rw = 1 To 12000 Step 1400
'        For J = 1 To 8000 Step 100
'            PosPrint I, J, "X"
'        Next J

        PosPrint 200, Rw, "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

    Next Rw
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub


Public Sub PrintForm99(ByVal FormType As String, ByVal TaxYear As Long, ByVal TestMode As Boolean)

    ' prglobal records used for 1096
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = 30 AND " & _
                "UserID = " & User.ID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.UserID = User.ID
        PRGlobal.TypeCode = 30
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    SQLString = " SELECT * FROM Form99 WHERE FormType = '" & FormType & "'" & _
                " AND TaxYear = " & TaxYear
    If Form99.GetBySQL(SQLString) = False Then
        MsgBox "Form NF: " & FormType & " " & TaxYear, vbExclamation
        End
    End If
    
    
    PrtInit ("Port")    ' "Port" = Portrait
    SetFont 10, Equate.Portrait
    Prvw.Caption = GLCompany.Name & " - 1099 Print " & TaxYear & " " & Form99.FormType
    
    ' ******************************
    Nudge = 30
    ' ******************************
    
    ' payer demographics
    ' skip blank fields
    J = 0
    For I = 1 To 5
        PayerDemo(I) = ""
        If I = 1 Then X = Trim(GLCompany.Name)
        If I = 2 Then X = Trim(GLCompany.Address1)
        If I = 3 Then X = Trim(GLCompany.Address2)
        If I = 4 Then X = Trim(GLCompany.Address3)
        If I = 5 Then X = Trim(GLCompany.CSZ)
        If X <> "" And X <> "0" Then
            J = J + 1
            PayerDemo(J) = X
        End If
    Next I
    
    FormCount = 0
    
    Dim Horz96X, Vert96X As Long
    
    If FormType = "1096" Then
        If TestMode = False Then
            PrintFormDetail 0, FormType, TaxYear, TestMode
        Else
            Dim T96 As Currency
            T96 = 1234567.89
            Form96_NumForms = 999
            Form96_FWT = T96
            Form96_TotalAmt = T96
            Form96_Final = "X"
            Form96_Title = "TITLE"
            Form96_Date = "DATE"
            Form96_Type = 1
            Form96_TaxYear = 2011
            Form96_NECX = "NNN"
            Form96_MiscX = "MMM"
            Form96_RX = "RRR"
            Form96_IntX = "III"
            Form96_DivX = "DDD"
            HorzNudge = 4
            VertNudge = 4
            PrintFormDetail 0, FormType, TaxYear, TestMode
        End If
    Else
        If TestMode = False Then
            SQLString = " SELECT * FROM Payee99 ORDER BY PayeeName"
            If Payee99.GetBySQL(SQLString) = False Then
                MsgBox "No Payee info found!", vbInformation
                GoBack
            End If
            Do
                PrintFormDetail Payee99.PayeeID, FormType, TaxYear, TestMode
                If Payee99.GetNext = False Then Exit Do
                If FormCount = Form99.FormsPerPg Then
                    FormFeed
                    FormCount = 0
                End If
            Loop
        Else
            HorzNudge = 4
            VertNudge = 4
            For I = 1 To Form99.FormsPerPg
                PrintFormDetail 0, FormType, TaxYear, TestMode
            Next I
        End If
    End If
    
    PrvwReturn = True
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub PrintFormDetail(ByVal PayeeID As Long, _
                            ByVal FormType As String, _
                            ByVal TaxYear As Long, _
                            ByVal TestMode As Boolean)
    
Dim TestString, TestAmt As String
Dim Form1096 As Boolean

    If FormType = "1096" Then
        Form1096 = True
    Else
        Form1096 = False
    End If

    ' dont' print if no detail
    If Form1096 = False And TestMode = False Then
        SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & PayeeID & _
                    " AND FormType = '" & FormType & "' " & _
                    " AND TaxYear = " & TaxYear
        If Detail99.GetBySQL(SQLString) = False Then Exit Sub
    End If

    TestString = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    TestAmt = "##,###,###.##"
    
    FormCount = FormCount + 1
    
    If FormCount = 1 Then vPos = Form99.FormVert1
    If FormCount = 2 Then vPos = Form99.FormVert2
    If FormCount = 3 Then vPos = Form99.FormVert3
    If FormCount = 4 Then vPos = Form99.FormVert4
    
    SQLString = " SELECT * FROM Field99 WHERE FormType = '" & FormType & "' " & _
                " AND TaxYear = " & TaxYear & _
                " ORDER BY FieldOrder"
    If Field99.GetBySQL(SQLString) = False Then
        MsgBox "No fields found for: " & FormType & " " & TaxYear, vbExclamation
        End
    End If
    
    Do
                
        ' test mode
        If PayeeID = 0 And Form1096 = False Then
            If Field99.FieldFormat = Equate.fmtAmount Then
                X = TestAmt
            Else
                X = Field99.BoxName
            End If
        Else
            X = "~"
            Select Case Field99.BoxName
                Case "TaxYear"
                    ' 2023-12-30 4 digit year
                    If TaxYear <= 2022 Then
                        X = TaxYear Mod 100
                    Else
                        X = TaxYear
                    End If
                Case "Payer1"
                    X = PayerDemo(1)
                Case "Payer2"
                    X = PayerDemo(2)
                Case "Payer3"
                    X = PayerDemo(3)
                Case "Payer4"
                    X = PayerDemo(4)
                Case "Payer5"
                    X = PayerDemo(5)
                Case "PayerFederalID"
                    X = GLCompany.FederalID
                Case "PayeeName"
                    X = Payee99.PayeeName
                Case "PayeeAddress"
                    X = Payee99.Address
                Case "PayeeCSZ"
                    X = Payee99.CSZ
                Case "PayeeAccountNumber"
                    X = Payee99.AccountNumber
                Case "PayeeFederalID"
                    X = Payee99.FederalID
                Case "SSN"
                    X = GLCompany.SSN
                Case "Contact"
                    X = PRGlobal.Var1
                Case "Email"
                    X = PRGlobal.Var2
                Case "Phone"
                    X = PRGlobal.Var3
                Case "Fax"
                    X = PRGlobal.Var4
                Case "Title"
                    X = PRGlobal.Var5
                Case "NumForms"
                    X = Form96_NumForms
                Case "FWT"
                    X = FormatAmt(Form96_FWT)
                Case "TotalAmt"
                    X = FormatAmt(Form96_TotalAmt)
                Case "Final"
                    X = Form96_Final
                Case "Title"
                    X = Form96_Title
                Case "Date"
                    X = Form96_Date
                Case "NECX"
                    X = Form96_NECX
                Case "MiscX"
                    X = Form96_MiscX
                Case "RX"
                    X = Form96_RX
                Case "IntX"
                    X = Form96_IntX
                Case "DivX"
                    X = Form96_DivX
                Case Else
                    SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & PayeeID & _
                                " AND FormType = '" & FormType & "' " & _
                                " AND TaxYear = " & TaxYear & _
                                " AND BoxName = '" & Field99.BoxName & "'"
                    If Detail99.GetBySQL(SQLString) = False Then
                        If Field99.FieldFormat = Equate.fmtAmount Then
                            X = FormatAmt("")
                        Else
                            X = ""
                        End If
                    Else
                        If Field99.FieldFormat = Equate.fmtAmount Then
                            X = FormatAmt(Detail99.FieldValue)
                        Else
                            X = Detail99.FieldValue
                        End If
                    End If
            
            End Select
            
            If X = "~" Then
                MsgBox "Field99.BoxName Invalid: " & Field99.BoxName, vbExclamation
                End
            End If
        
        End If
        
        PosPrint Field99.HorzPosn, vPos + Field99.VertPosn, X
        
        If Field99.GetNext = False Then Exit Do
    
    Loop
    
End Sub

Public Function FormatAmt(ByVal AmtString) As String

Dim Amt99 As Currency

    On Error Resume Next
    Amt99 = AmtString
    If Err.Number <> 0 Then
        Amt99 = 0
    End If
    On Error GoTo 0
    X = Format(Amt99, "##,###,##0.00")
    X = Trim(X)
    FormatAmt = Space(12 - Len(X)) & X

End Function
