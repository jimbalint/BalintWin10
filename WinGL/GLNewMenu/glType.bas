Attribute VB_Name = "glType"
Option Explicit

Public Function glTypeByte(ByVal AcctType As String) As Integer
    glTypeByte = 0
    If AcctType = "0" Then glTypeByte = 1
    If AcctType = "T" Then glTypeByte = 2
    If AcctType = "H" Then glTypeByte = 3
    If AcctType = "D" Then glTypeByte = 4
    If AcctType = "I" Then glTypeByte = 5
    If AcctType = "L" Then glTypeByte = 6
    If AcctType = "A" Then glTypeByte = 7
    If AcctType = "E" Then glTypeByte = 8
    If AcctType = "U" Then glTypeByte = 9
    If AcctType = "S" Then glTypeByte = 10
    If AcctType = "." Then glTypeByte = 11
    If AcctType = "M" Then glTypeByte = 12
    If AcctType = "B" Then glTypeByte = 13
    If AcctType = "P" Then glTypeByte = 14
    If AcctType = "C" Then glTypeByte = 15
End Function

Public Function glTypeChar(ByVal ndx As Byte) As String

    glTypeChar = " "
    If ndx = 1 Then glTypeChar = "0"
    If ndx = 2 Then glTypeChar = "T"
    If ndx = 3 Then glTypeChar = "H"
    If ndx = 4 Then glTypeChar = "D"
    If ndx = 5 Then glTypeChar = "I"
    If ndx = 6 Then glTypeChar = "L"
    If ndx = 7 Then glTypeChar = "A"
    If ndx = 8 Then glTypeChar = "E"
    If ndx = 9 Then glTypeChar = "U"
    If ndx = 10 Then glTypeChar = "S"
    If ndx = 11 Then glTypeChar = "."
    If ndx = 12 Then glTypeChar = "M"
    If ndx = 13 Then glTypeChar = "B"
    If ndx = 14 Then glTypeChar = "P"
    If ndx = 15 Then glTypeChar = "C"

End Function

Public Function glTypeName(ByVal ndx As Byte) As String

    glTypeName = "ERROR"
    If ndx = 0 Then glTypeName = "BLANK"
    If ndx = 1 Then glTypeName = "ZERO ACCOUNT POSTABLE"
    If ndx = 2 Then glTypeName = "TOTAL RECORD"
    If ndx = 3 Then glTypeName = "HEADING OR DESCRIPTIVE RECORD"
    If ndx = 4 Then glTypeName = "DATE ROUTINE"
    If ndx = 5 Then glTypeName = "INCOME CATEGORY"
    If ndx = 6 Then glTypeName = "LIABILITY OR CAPITAL CATEGORY"
    If ndx = 7 Then glTypeName = "ASSET CATEGORY"
    If ndx = 8 Then glTypeName = "EXPENSE CATEGORY"
    If ndx = 9 Then glTypeName = "UNDERLINE RECORD"
    If ndx = 10 Then glTypeName = "SIGN RECORD"
    If ndx = 11 Then glTypeName = "PERCENT BASE"
    If ndx = 12 Then glTypeName = "MATH RECORD"
    If ndx = 13 Then glTypeName = "BALANCE SHEET"
    If ndx = 14 Then glTypeName = "PROFIT AND LOSS"
    If ndx = 15 Then glTypeName = "CLEARING RECORD"

End Function


