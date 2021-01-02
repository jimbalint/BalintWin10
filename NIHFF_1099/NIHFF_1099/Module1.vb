Option Explicit On

Imports System.IO
Imports System.Text

Module Module1

    Dim TaxYear As Integer = 2020

    Dim InputFolder As String = "C:\aSend\NIHFF_20"
    Dim UploadFolder As String = InputFolder & "\UploadFiles"
    Dim dtCompany As DataTable
    Dim dtFiles As DataTable
    Dim dtForms As DataTable

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

    Sub ProcessFiles()
        For Each rw As DataRow In dtFiles.Rows
            Console.WriteLine(rw.Item("FileName") & vbTab & rw.Item("FormType"))
            If rw.Item("FormType") = "MISC" Then
                ProcessMiscFile(rw)
            Else
                Console.WriteLine("1099-NEC")
            End If

            x = UploadFolder & "\" & Replace(rw.Item("FileName"), ".txt", "-upl.txt")
            Dim sw As New StreamWriter(x)
            tRecord(sw)
            sw.Close()

            Console.WriteLine("==============")
        Next
    End Sub

    Sub tRecord(ByRef sw As StreamWriter)

        'Type            string(1)
        'PayYear         string(4)
        'PriorYear       string(1)
        'TIN             string(9)
        'CCode           string(5)
        'b1              string(7)
        'TestFile        string(1)
        'Foreign         string(1)
        'Name            string(40)
        'Name2           string(40)
        'CompName        string(40)
        'CompName2       string(40)
        'CompAddr        string(40)
        'CompCity        string(40)
        'CompState       string(2)
        'CompZip         string(9)
        'b2              string(15)
        'PayeeCt         string(8)
        'ContactName     string(40)
        'ContactPhone    string(15)
        'ContactEMail    string(50)
        'Tape            string(2)
        'MediaNum        string(6)
        'b3              string(83)
        'SeqNumber       string(8)
        'b4              string(10)
        'VendInd         string(1)
        'VendName        string(40)
        'VendAddr        string(40)
        'VendCity        string(40)
        'VendState       string(2)
        'VendZip         string(9)
        'VendContact     string(40)
        'VendPhone       string(15)
        'b5              string(35)
        'VendForeign     string(1)
        'b6              string(8)
        'b7              string(2)


        sw.Write(FixedLen("AAA", 5))
        sw.Write(FixedLen("BBB", 10))
        sw.Write(FixedLen("CCC", 15))
        sw.Write(vbCrLf)
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

    Sub ProcessMiscFile(ByVal rw1 As DataRow)

        i = 0
        j = 0
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

    End Sub

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

    Sub Init()

        DefineDT()
        AddCompanyInfo()

        x = Dir(InputFolder & "\*.*")
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

        rw = dtCompany.NewRow

        rw("ContactName") = "Rebecca Foldi"
        rw("ContactPhone") = "330-849-6926"
        rw("ContactEMail") = "RFoldi@INVENT.ORG"
        rw("CCode") = "17677"
        rw("TTIN") = "341580038"
        rw("TName1") = "National Inventors Hall of Fame Foundation, Inc."
        rw("TName2") = ""
        rw("TCompName1") = "National Inventors Hall of Fame Foundation, Inc."
        rw("TCompName2") = ""
        rw("TCompAddr") = "221 S. Broadway"
        rw("TCompCity") = "Akron"
        rw("TCompState") = "OH"
        rw("TCompZip") = "44308"

        dtCompany.Rows.Add(rw)

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

        dtCompany.Columns.Add("ContactName")
        dtCompany.Columns.Add("ContactPhone")
        dtCompany.Columns.Add("ContactEMail")
        dtCompany.Columns.Add("CCode")
        dtCompany.Columns.Add("TTIN")
        dtCompany.Columns.Add("TName1")
        dtCompany.Columns.Add("TName2")
        dtCompany.Columns.Add("TCompName1")
        dtCompany.Columns.Add("TCompName2")
        dtCompany.Columns.Add("TCompAddr")
        dtCompany.Columns.Add("TCompCity")
        dtCompany.Columns.Add("TCompState")
        dtCompany.Columns.Add("TCompZip")

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
