Imports System.IO

Public Class clsExport

    Public sw As StreamWriter
    Public dct As Dictionary(Of String, Object)
    Public dctCompany As Dictionary(Of String, String)
    Public TaxYear As Integer
    Public tTest As String
    Public Corrected As String
    Public ReturnType As String
    Public SeqNum As Integer
    Dim Total(16) As Double

    Public Sub tRecord()

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
        If tTest <> "T" Then
            sw.Write(tTest)
        Else
            sw.Write(" ")
        End If
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


    Public Sub aRecord()
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
        sw.Write(FixedLen(ReturnType, 2))
        '! AmountCodes     string(14)   ! 3 = Other Income   4 = FWT
        'AmountCodes     string(16)   ! 3 = Other Income   4 = FWT  - 2011 expanded 2 positions - 2013 - added "9"
        '                            ! 2015 - removed "9"
        If ReturnType = "A" Then
            sw.Write(FixedLen("12345678ABCDE", 16))
        Else
            sw.Write(FixedLen("14", 16))
        End If

        '! b2              string(10) - 2011 changed from 10 to 8
        'b2              string(8)
        sw.Write(StrDup(8, " "))
        'Foreign         string(1)
        sw.Write(StrDup(1, " "))
        'Name1           string(40)
        sw.Write(FixedLen(dctCompany("TName1"), 40))
        'Name2           string(40)
        sw.Write(FixedLen(dctCompany("TName2"), 40))
        'XferAgent       string(1)
        sw.Write("0")
        'ShipAddr        string(40)
        sw.Write(FixedLen(dctCompany("TCompAddr"), 40))
        'City            string(40)
        sw.Write(FixedLen(dctCompany("TCompCity"), 40))
        'State           string(2)
        sw.Write(FixedLen(dctCompany("TCompState"), 2))
        'Zip             string(9)
        sw.Write(FixedLen(dctCompany("TCompZip"), 9))
        'Phone           string(15)
        sw.Write(FixedLen(dctCompany("ContactPhone"), 15))
        'b3              string(260)
        sw.Write(StrDup(260, " "))
        'SeqNumber       string(8)
        sw.Write(IntString(2, 8))
        'b4              string(241)
        sw.Write(StrDup(241, " "))
        'b5              string(2)
        sw.WriteLine(StrDup(2, " "))

    End Sub

    Public Sub bRecords(ByRef dt As DataTable)

        Dim AcctNum As Integer = 0
        SeqNum = 2

        For Each rw As DataRow In dt.Rows
            'Type            string(1)
            sw.Write("B")
            'PayYear         string(4)
            sw.Write(TaxYear)
            'Corrected       string(1)
            sw.Write(FixedLen(Corrected, 1))
            'NameControl     string(4)   !!!!!!!!!!!!!!!!!
            sw.Write(StrDup(4, " "))

            'TINType         string(1)   ! 1 = EIN  2 = SSN
            'if tps:PayeeID = '' or len(tps:PayeeID) < 3
            '    b:TINType = ''
            'elsif sub(tps:PayeeID,3,1) = '-'        ! EIN
            '    b:TINType = '1'
            'elsif sub(tps:PayeeID,4,1) = '-'        ! SSN
            '    b:TINType = '2'
            '    Else
            '    b:TINType = ''
            '        End
            Dim pid As String = Trim(rw("PayeeID"))
            If pid = "" Or Len(pid) < 3 Then
                sw.Write(" ")
            ElseIf Mid(pid, 3, 1) = "-" Then
                sw.Write("1")
            ElseIf Mid(pid, 4, 1) = "-" Then
                sw.Write("2")
            Else
                sw.Write(" ")
            End If

            'TIN             string(9)
            sw.Write(FixedLen(DigitsOnly(pid), 9))

            'AcctNum         string(20)
            '! changed in tax year 2009
            '!   use unique number
            AcctNum += 1
            sw.Write(IntString(AcctNum, 20))

            'OfficeCode      string(4)
            sw.Write(StrDup(4, " "))
            'b1              string(10)
            sw.Write(StrDup(10, " "))

            Dim pay(16) As Double
            Dim ii As Integer
            For ii = 1 To 16
                pay(ii) = 0
                Total(ii) = 0
            Next

            If ReturnType = "A" Then
                ' MISC
                Select Case rw("AmountLine")
                    Case 1
                        pay(1) = rw("Amount")
                        Total(1) += rw("Amount")
                    Case 6
                        pay(3) = rw("Amount")
                        Total(3) += rw("Amount")
                    Case 9
                        pay(7) = rw("Amount")
                        Total(7) += rw("Amount")
                    Case Else
                        MsgBox("Bad Amount Line: " & rw("AmountLine") & " " & "Payee: " & rw("PayeeName"))
                        End
                End Select
            Else
                ' NEC
                Select Case rw("AmountLine")
                    Case 1
                        pay(1) = rw("Amount")
                        Total(1) += rw("Amount")
                    Case Else
                        MsgBox("Bad Amount Line: " & rw("AmountLine") & " " & "Payee: " & rw("PayeeName"))
                        End
                End Select
            End If

            For ii = 1 To 16
                sw.Write(AmtString(pay(ii), 12))
            Next

            'Foreign         string(1)
            sw.Write(" ")
            'Name1           string(40)
            sw.Write(FixedLen(rw("PayeeName"), 40))
            'Name2           string(40)
            sw.Write(StrDup(40, " "))
            'b3              string(40)
            sw.Write(StrDup(40, " "))
            'Addr1           string(40)
            sw.Write(FixedLen(rw("PayeeAddr"), 40))
            'b4              string(40)
            sw.Write(StrDup(40, " "))
            'City            string(40)
            sw.Write(FixedLen(rw("PayeeCity"), 40))
            'State           string(2)
            sw.Write(FixedLen(rw("PayeeState"), 2))
            'Zip             string(9)
            sw.Write(FixedLen(rw("PayeeZip"), 9))
            'b5              string(1)
            sw.Write(" ")
            'SeqNumber       string(8)
            SeqNum += 1
            sw.Write(IntString(SeqNum, 8))
            'b6              string(36)
            sw.Write(StrDup(36, " "))

            If ReturnType = "A" Then
                ' MISC
                'TIN2            string(1)
                sw.Write(" ")
                'b7              string(2)
                sw.Write(StrDup(2, " "))
                'DirectSales     string(1)
                sw.Write(" ")
                'b8              string(115)
                sw.Write(StrDup(115, " "))
                'SpecialData     string(60)
                sw.Write(StrDup(60, " "))
                'SWT             string(12)
                sw.Write(StrDup(12, "0"))
                'CWT             string(12)
                sw.Write(StrDup(12, "0"))
                'CombCode        string(2)
                sw.Write(StrDup(2, " "))
                'b9              string(2)
                sw.WriteLine(StrDup(2, " "))
            Else
                ' NEC
                ' 2nd TIN notice
                sw.Write(" ")
                ' blank
                sw.Write(StrDup(3, " "))
                ' FATCA
                sw.Write(" ")
                ' blank
                sw.WriteLine(StrDup(202, " "))
            End If

        Next

    End Sub

    Public Sub cRecord()
        'Type            string(1)
        sw.Write("C")
        'Count           string(8)
        sw.Write(IntString(SeqNum - 2, 8))
        'b1              string(6)
        sw.Write(StrDup(6, " "))

        Dim ii As Integer
        For ii = 1 To 16
            sw.Write(AmtString(Total(ii), 18))
        Next

        sw.Write(StrDup(196, " "))
        'SeqNumber       string(8)
        SeqNum += 1
        sw.Write(IntString(SeqNum, 8))
        'b3              string(241)
        sw.Write(StrDup(241, " "))
        'b4              string(2)
        sw.WriteLine(StrDup(2, " "))

    End Sub

    Public Sub fRecord()
        sw.Write("F")
        sw.Write(IntString(1, 8))
        sw.Write(StrDup(21, "0"))
        sw.Write(StrDup(19, " "))
        sw.Write(IntString(SeqNum - 3, 8))
        sw.Write(StrDup(442, " "))
        SeqNum += 1
        sw.Write(IntString(SeqNum, 8))
        sw.Write(StrDup(241, " "))
        sw.WriteLine(StrDup(2, " "))
    End Sub

End Class
