﻿Option Explicit On

Imports System.IO
Imports System.Text

Module Module1

    Dim TaxYear As Integer = 2020
    Dim tTest As String = "T"
    Dim Corrected As String = " "

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

        Dim exp As New clsExport
        exp.TaxYear = TaxYear
        exp.tTest = tTest
        exp.dctCompany = dctCompany
        exp.Corrected = Corrected

        For Each rw As DataRow In dtFiles.Rows

            Dim ReturnType As String
            Console.WriteLine(rw.Item("FileName") & vbTab & rw.Item("FormType"))
            If rw.Item("FormType") = "MISC" Then
                dct = ProcessMiscFile(rw)
                ReturnType = "A"
            Else
                ReturnType = "NE"
                Console.WriteLine("1099-NEC")
            End If

            x = UploadFolder & "\" & Replace(rw.Item("FileName"), ".txt", "-upl.txt")
            Dim sw As New StreamWriter(x)

            exp.sw = sw
            exp.dct = dct
            exp.ReturnType = ReturnType
            exp.tRecord()
            exp.aRecord()
            exp.bRecords(dtForms)
            exp.cRecord()
            exp.fRecord()

            sw.Close()

            ' add routine to check for each line is 750 chars

            Console.WriteLine("==============")

            dtForms.Clear()

        Next
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

    Sub ProcessNECFile(ByVal fnm As String)

    End Sub


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


End Module
