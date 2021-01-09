Option Explicit On

Imports System.IO
Imports System.Text

Module Module1

    Dim TaxYear As Integer = 2020
    Dim tTest As String = "T"

    Dim InputFolder As String = "C:\aSend\NIHFF_20"
    Dim UploadFolder As String = InputFolder & "\UploadFiles"
    Dim dctCompany As New Dictionary(Of String, String)
    Dim dtFiles As DataTable
    Dim dtForms As DataTable
    Dim SeqNumber As Integer = 0

    Dim x, y, z As String
    Dim i, j, k As Integer
    Dim rw As DataRow

    Sub Main()
        Init()
        ProcessFiles()
        DebugOutput("C:\aSend\NIHFF_20\Debug\NIHFF.txt")
        Console.WriteLine("...")
        Console.ReadKey()
    End Sub

    Function Test1()
        Dim d As New Dictionary(Of String, Double)
        d.Add("a", 10)
        d.Add("b", 20)
        Return d
    End Function

    Sub ProcessFiles()
        Dim dct As New Dictionary(Of String, Object)
        For Each rw As DataRow In dtFiles.Rows
            Console.WriteLine(rw.Item("FileName") & vbTab & rw.Item("FormType"))
            If rw.Item("FormType") = "MISC" Then
                dct = ProcessMiscFile(rw)
            Else
                Console.WriteLine("1099-NEC")
            End If

            x = UploadFolder & "\" & Replace(rw.Item("FileName"), ".txt", "-upl.txt")
            Dim sw As New StreamWriter(x)
            tRecord(sw, dct)
            sw.Close()

            Console.WriteLine("==============")
        Next
    End Sub

    Sub aRecord(ByRef sw As StreamWriter, ByVal dct As Dictionary(Of String, Object))
        'Payer           group,over(out:Record),pre(A)
        'Type            string(1)
        sw.Write("A")
        'PayYear         string(4)
        sw.Write(TaxYear)
        'b1              string(6)
        sw.Write(StrDup(6, " "))
        'TIN             string(9)
        sw.Write(dctCompany("TTIN"))
        'NameControl     string(4)
        sw.Write(StrDup(4, " "))
        'LastFiling      string(1)
        sw.Write(StrDup(1, " "))
        'ReturnType      string(2)    ! A
        sw.Write("A")
        '! AmountCodes     string(14)   ! 3 = Other Income   4 = FWT
        'AmountCodes     string(16)   ! 3 = Other Income   4 = FWT  - 2011 expanded 2 positions - 2013 - added "9"
        '                            ! 2015 - removed "9"
        FixedLen("12345678ABCDE", 16)
        '! b2              string(10) - 2011 changed from 10 to 8
        'b2              string(8)
        'Foreign         string(1)
        'Name1           string(40)
        'Name2           string(40)
        'XferAgent       string(1)
        'ShipAddr        string(40)
        'City            string(40)
        'State           string(2)
        'Zip             string(9)
        'Phone           string(15)
        'b3              string(260)
        'SeqNumber       string(8)
        'b4              string(241)
        'b5              string(2)

    End Sub

    Sub tRecord(ByRef sw As StreamWriter, ByVal dct As Dictionary(Of String, Object))

        'Type            string(1)
        sw.Write("T")
        'PayYear         string(4)
        sw.Write(TaxYear)
        'PriorYear       string(1)
        sw.Write(" ")
        'TIN             string(9)
        sw.Write(dctCompany("TTIN"))
        'CCode           string(5)
        sw.Write(dctCompany("CCode"))
        'b1              string(7)
        sw.Write(StrDup(7, " "))
        'TestFile        string(1)
        sw.Write(tTest)
        'Foreign         string(1)
        sw.Write(" ")
        'Name            string(40)
        sw.Write(FixedLen(dctCompany("TName1"), 40))
        'Name2           string(40)
        sw.Write(FixedLen(dctCompany("TName2"), 40))
        'CompName        string(40)
        sw.Write(FixedLen(dctCompany("TCompName1"), 40))
        'CompName2       string(40)
        sw.Write(FixedLen(dctCompany("TCompName2"), 40))
        'CompAddr        string(40)
        sw.Write(FixedLen(dctCompany("TCompAddr"), 40))
        'CompCity        string(40)
        sw.Write(FixedLen(dctCompany("TCompCity"), 40))
        'CompState       string(2)
        sw.Write(FixedLen(dctCompany("TCompState"), 2))
        'CompZip         string(9)
        sw.Write(FixedLen(dctCompany("TCompZip"), 9))
        'b2              string(15)
        sw.Write(StrDup(15, " "))
        'PayeeCt         string(8)
        sw.Write(IntString(dct("Count"), 8))
        'ContactName     string(40)
        sw.Write(FixedLen(dctCompany("ContactName"), 40))
        'ContactPhone    string(15)
        sw.Write(FixedLen(dctCompany("ContactPhone"), 15))
        'ContactEMail    string(50)
        sw.Write(FixedLen(dctCompany("ContactEMail"), 50))
        'Tape            string(2)
        'MediaNum        string(6)
        'b3              string(83)
        sw.Write(StrDup(91, " "))
        'SeqNumber       string(8)
        sw.Write(IntString(1, 8))
        'b4              string(10)
        sw.Write(StrDup(10, " "))
        'VendInd         string(1)
        sw.Write("I")
        'VendName        string(40)
        sw.Write(StrDup(40, " "))
        'VendAddr        string(40)
        sw.Write(StrDup(40, " "))
        'VendCity        string(40)
        sw.Write(StrDup(40, " "))
        'VendState       string(2)
        sw.Write(StrDup(2, " "))
        'VendZip         string(9)
        sw.Write(StrDup(9, " "))
        'VendContact     string(40)
        sw.Write(StrDup(40, " "))
        'VendPhone       string(15)
        sw.Write(StrDup(15, " "))
        'b5              string(35)
        sw.Write(StrDup(35, " "))
        'VendForeign     string(1)
        sw.Write(StrDup(1, " "))
        'b6              string(8)
        sw.Write(StrDup(8, " "))
        'b7              string(2)
        sw.WriteLine(StrDup(2, " "))

    End Sub

    Sub DebugOutput(ByVal fnm As String)
        If fnm = "" Then Exit Sub
        Dim sw As New StreamWriter(fnm)
        For Each rw As DataRow In dtForms.Rows
            For Each fld As DataColumn In dtForms.Columns
                x = fld.ColumnName & vbTab & rw(fld.ColumnName)
                sw.WriteLine(x)
            Next
            sw.WriteLine("---------------------------------")
        Next
        sw.Close()
    End Sub

    Function ProcessMiscFile(ByVal rw1 As DataRow)

        i = 0
        j = 0

        Dim dPay01 As Double = 0
        Dim dPay03 As Double = 0
        Dim dPay07 As Double = 0

        rw = dtForms.NewRow
        rw("FileName") = rw1("FileName")

        z = InputFolder & "\" & rw1.Item("FileName")
        Dim sr As New StreamReader(z)
        Do While Not sr.EndOfStream
            y = sr.ReadLine
            i += 1
            If i = 15 Then
                j += 1
                i = 1
                If j Mod 10 = 1 Then Console.WriteLine(rw("FileName") & vbTab & j)

                dtForms.Rows.Add(rw)
                rw = dtForms.NewRow
                rw("FileName") = rw1("FileName")

            End If

            Select Case i
                Case 1
                    rw("Payer1") = Mid(y, 6, 37)
                    If Trim(Mid(y, 43, 10)) <> "" Then
                        rw("Amount") = CDbl(Mid(y, 43, 10))
                        dPay01 += CDbl(Mid(y, 43, 10))
                        rw("AmountLine") = 1
                        rw("Box") = "Box #1 Rents"
                    End If

                Case 2
                    rw("Payer2") = Mid(y, 6, 50)
                Case 3
                    rw("Payer3") = Mid(y, 6, 50)
                Case 4
                    rw("PayerCity") = Mid(y, 6, 20)
                    rw("PayerState") = Mid(y, 27, 2)
                    rw("PayerZip") = ZipString(Mid(y, 30, 10))
                Case 5
                    rw("PayerPhone") = Mid(y, 6, 50)
                Case 6
                    If Trim(Mid(y, 43, 10)) <> "" Then
                        rw("Amount") = CDbl(Mid(y, 43, 10))
                        dPay03 += CDbl(Mid(y, 43, 10))
                        rw("AmountLine") = 6
                        rw("Box") = "Box #3 Other Income"
                    End If
                Case 7
                    rw("FID") = Mid(y, 6, 10)
                    rw("FID2") = DigitsOnly(rw("FID"))
                    rw("PayeeID") = Mid(y, 23, 20)
                    rw("PayeeID2") = DigitsOnly(rw("PayeeID"))
                Case 8
                    rw("PayeeName") = Mid(y, 6, 50)
                Case 9
                    If Trim(Mid(y, 43, 10)) <> "" Then
                        rw("Amount") = CDbl(Mid(y, 43, 10))
                        dPay07 += CDbl(Mid(y, 43, 10))
                        rw("AmountLine") = 9
                        rw("Box") = "Box #7 Non Emp Comp"
                    End If
                Case 11
                    rw("PayeeAddr") = Mid(y, 6, 50)
                Case 12
                    rw("PayeeCity") = Mid(y, 6, 21)
                    rw("PayeeState") = Mid(y, 27, 2)
                    rw("PayeeZip") = ZipString(Mid(y, 30, 10))

            End Select

        Loop
        sr.Close()

        j += 1

        dtForms.Rows.Add(rw)
        Console.WriteLine(rw("FileName") & vbTab & j)

        Dim dct As New Dictionary(Of String, Object)
        dct.Add("Count", j)
        dct.Add("Pay01", dPay01)
        dct.Add("Pay03", dPay03)
        dct.Add("Pay07", dPay07)
        Return (dct)

    End Function

    Sub processNECFile(ByVal fnm As String)

    End Sub

    Function FixedLen(ByVal str As String, ByVal slen As Integer) As String
        str = Trim(str)
        If slen < Len(str) Then
            FixedLen = Left(str, slen)
        Else
            FixedLen = str & StrDup(slen - Len(str), " ")
        End If
    End Function

    Function IntString(ByVal str As String, ByVal slen As Integer) As String
        Dim iint As Integer = CInt(str)
        IntString = iint.ToString("D" & slen)
    End Function
    Function AmtString(ByVal str As String, ByVal slen As Integer) As String
        Dim dbl As Double = CDbl(str) * 100
        AmtString = dbl.ToString("D" & slen)
    End Function

    Sub Init()

        DefineDT()
        AddCompanyInfo()

        x = Dir(InputFolder & "\*.txt")
        Do While x > ""
            rw = dtFiles.NewRow
            rw("FileName") = x
            rw("FormType") = IIf(InStr(x, "MISC", CompareMethod.Text), "MISC", "NEC")
            rw("FormCount") = 0
            rw("TotalAmount") = 0
            dtFiles.Rows.Add(rw)
            x = Dir()
        Loop

        If Not (Directory.Exists(UploadFolder)) Then
            Directory.CreateDirectory(UploadFolder)
        Else
            x = Dir(UploadFolder & "\*.*")
            Do While x <> ""
                Kill(UploadFolder & "\" & x)
                x = Dir()
            Loop
        End If

    End Sub

    Sub AddCompanyInfo()
        dctCompany.Add("ContactName", "Rebecca Foldi")
        dctCompany.Add("ContactPhone", "330-849-6926")
        dctCompany.Add("ContactEMail", "RFoldi@INVENT.ORG")
        dctCompany.Add("CCode", "17677")
        dctCompany.Add("TTIN", "341580038")
        dctCompany.Add("TName1", "National Inventors Hall of Fame Foundation, Inc.")
        dctCompany.Add("TName2", "")
        dctCompany.Add("TCompName1", "National Inventors Hall of Fame Foundation, Inc.")
        dctCompany.Add("TCompName2", "")
        dctCompany.Add("TCompAddr", "221 S. Broadway")
        dctCompany.Add("TCompCity", "Akron")
        dctCompany.Add("TCompState", "OH")
        dctCompany.Add("TCompZip", "44308")
    End Sub

    Sub DefineDT()

        dtFiles = New DataTable("Files")
        dtFiles.Columns.Add("FileName")
        dtFiles.Columns.Add("FormType")
        dtFiles.Columns.Add("FormCount")
        dtFiles.Columns.Add("TotalAmount")

        dtForms = New DataTable("Forms")
        dtForms.Columns.Add("FileName")
        dtForms.Columns.Add("FID")
        dtForms.Columns.Add("FID2")
        dtForms.Columns.Add("NameID")
        dtForms.Columns.Add("Payer1")
        dtForms.Columns.Add("Payer2")
        dtForms.Columns.Add("Payer3")
        dtForms.Columns.Add("PayerCity")
        dtForms.Columns.Add("PayerState")
        dtForms.Columns.Add("PayerZip")
        dtForms.Columns.Add("PayerPhone")
        dtForms.Columns.Add("PayeeID")
        dtForms.Columns.Add("PayeeID2")
        dtForms.Columns.Add("AmountLine")
        dtForms.Columns.Add("Box")
        dtForms.Columns.Add("PayeeName")
        dtForms.Columns.Add("Amount")
        dtForms.Columns.Add("PayeeAddr")
        dtForms.Columns.Add("PayeeCity")
        dtForms.Columns.Add("PayeeState")
        dtForms.Columns.Add("PayeeZip")

    End Sub

    Function ZipString(ByVal InString As String) As String
        InString = Trim(InString)
        If Len(Trim(InString)) <= 5 Then Return InString
        Return (InString.Replace("-", ""))
    End Function

    Function DigitsOnly(ByVal InString As String) As String
        InString = Trim(InString)
        DigitsOnly = ""
        For ii As Integer = 1 To Len(InString)
            If InStr("0123456789", Mid(InString, ii, 1), CompareMethod.Text) Then
                DigitsOnly &= Mid(InString, ii, 1)
            End If
        Next
    End Function
End Module
