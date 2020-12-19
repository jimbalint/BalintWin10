Option Explicit On

Imports System.IO
Imports System.Text

Module Module1

    Dim TaxYear As Integer = 2020

    Dim InputFolder As String = "C:\aSend\NIHFF_20"
    Dim UploadFolder As String = InputFolder & "\UploadFiles"
    Dim dtFiles As DataTable
    Dim dtForms As DataTable

    Dim ContactName As String = "Rebecca Foldi"
    Dim ContactPhone As String = "330-849-6926"
    Dim ContactEMail As String = "RFoldi@INVENT.ORG"
    Dim CCode As String = "17677"
    Dim TTIN As String = "341580038"

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

            Console.WriteLine("==============")
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


    Sub Init()

        DefineDT()

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
