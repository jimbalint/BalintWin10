Attribute VB_Name = "modRC4"
Public Function RC4Encrypt(ByVal text As String, ByVal encryptkey As String)

    If text = "" Then
        RC4Encrypt = ""
        Exit Function
    End If

    Dim sbox(256)
    Dim Key(256)
    Dim Temp As Integer
    Dim a As Long
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim cipherby As Integer
    Dim cipher As String
    I = 0
    J = 0
    RC4Initialize encryptkey, Key, sbox
    For a = 1 To Len(text)
        I = (I + 1) Mod 256
        J = (J + sbox(I)) Mod 256
        Temp = sbox(I)
        sbox(I) = sbox(J)
        sbox(J) = Temp
        K = sbox((sbox(I) + sbox(J)) Mod 256)
        cipherby = (Asc(Mid$(text, a, 1))) Xor K
        If Len(Hex(cipherby)) = 1 Then
            cipher = cipher & "0" & Hex(cipherby)
        Else
            cipher = cipher & Hex(cipherby)
        End If
    Next
    RC4Encrypt = cipher
End Function

 Public Function RC4Decrypt(ByVal text As String, ByVal encryptkey As String)
 
    If text = "" Then
        RC4Decrypt = ""
        Exit Function
    End If
 
    Dim sbox(256) As Integer
    Dim Key(256) As Integer
    Dim Text2 As String
    Dim Temp As Integer
    Dim a As Long
    Dim I As Integer
    Dim J As Integer
    Dim K As Long
    Dim w As Integer
    Dim cipherby As Integer
    Dim cipher As String
    For w = 1 To Len(text) Step 2
        Text2 = Text2 & Chr(Dec(Mid$(text, w, 2)))
    Next
    I = 0
    J = 0
    RC4Initialize encryptkey, Key, sbox
    For a = 1 To Len(Text2)
        I = (I + 1) Mod 256
        J = (J + sbox(I)) Mod 256
        Temp = sbox(I)
        sbox(I) = sbox(J)
        sbox(J) = Temp
        K = sbox((sbox(I) + sbox(J)) Mod 256)
        cipherby = Asc(Mid$(Text2, a, 1)) Xor K
        cipher = cipher & Chr(cipherby)
    Next
    RC4Decrypt = cipher
End Function

Public Function RC4Initialize(strPwd, ByRef Key, ByRef sbox)
    Dim tempSwap
    Dim a
    Dim b
    Dim intlength As Long
    intlength = Len(strPwd)
    For a = 0 To 255
        Key(a) = Asc(Mid$(strPwd, a Mod intlength + 1, 1))
        sbox(a) = a
    Next
    b = 0
    For a = 0 To 255
        b = (b + sbox(a) + Key(a)) Mod 256
        tempSwap = sbox(a)
        sbox(a) = sbox(b)
        sbox(b) = tempSwap
    Next
End Function
 
Public Function Dec(Number) As String
    Dim base As String
    Dim iLen As Integer
    Dim iReturn As Long
    Dim I As Long
    Dim iTemp As String
    base = "0123456789ABCDEF"
    iLen = Len(Number)
    For I = 0 To iLen - 1
        iTemp = Mid$(Number, iLen - I, 1)
        iReturn = iReturn + (16 ^ I) * (InStr(1, base, iTemp) - 1)
    Next
    Dec = iReturn
End Function
