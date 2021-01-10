Module modGlobal

    Public Function ZipString(ByVal InString As String) As String
        InString = Trim(InString)
        If Len(Trim(InString)) <= 5 Then Return InString
        Return (InString.Replace("-", ""))
    End Function

    Public Function DigitsOnly(ByVal InString As String) As String
        InString = Trim(InString)
        DigitsOnly = ""
        For ii As Integer = 1 To Len(InString)
            If InStr("0123456789", Mid(InString, ii, 1), CompareMethod.Text) Then
                DigitsOnly &= Mid(InString, ii, 1)
            End If
        Next
    End Function

    Public Function FixedLen(ByVal str As String, ByVal slen As Integer) As String
        str = Trim(str)
        If slen < Len(str) Then
            FixedLen = Left(str, slen)
        Else
            FixedLen = str & StrDup(slen - Len(str), " ")
        End If
    End Function

    Public Function IntString(ByVal str As String, ByVal slen As Integer) As String
        Dim iint As Integer = CInt(str)
        IntString = iint.ToString("D" & slen)
    End Function

    Public Function AmtString(ByVal str As String, ByVal slen As Integer) As String
        Dim dbl As Double = CDbl(str) * 100
        AmtString = Right(StrDup(slen, "0") & dbl.ToString, slen)
    End Function

End Module
