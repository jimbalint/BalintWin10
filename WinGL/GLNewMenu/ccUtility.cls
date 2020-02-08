VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function UserName(ByVal ID As Long) As String
    On Error GoTo glErr
    UserName = "USER"
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM users WHERE ID=" & ID)
    UserName = rs!name
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Function ShowValue(ByVal Amount As Currency) As String
    ShowValue = FormatCurrency(Amount, 2)
End Function

Public Function ShowDate(ByVal thisDate As Date) As String
    ShowDate = Format(thisDate, "mm/dd/yyyy")
End Function

Private Sub txtSource_GotFocus()    ' Each Box (Field) sets up this way
    cmdSave.Enabled = True          ' Save on Edit
    txtSource.SelStart = 0          ' select data on entry
    txtSource.SelLength = Len(txtSource.Text)
End Sub

Public Sub SetListView(ByRef lv As ListView)
    lv.FullRowSelect = True
    lv.GridLines = True
    lv.LabelEdit = lvwManual
    lv.View = lvwReport
    lv.HideSelection = False
End Sub

Public Function GetCommandLine(Optional MaxArgs)
    Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
    If IsMissing(MaxArgs) Then MaxArgs = 10
    ReDim argarray(MaxArgs)
    NumArgs = 0
    InArg = False
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    For I = 1 To CmdLnLen
        C = Mid(CmdLine, I, 1)
        If (C <> " " And C <> vbTab) Then
            If Not InArg Then
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                InArg = True
            End If
            argarray(NumArgs) = argarray(NumArgs) & C
        Else
            InArg = False
        End If
    Next I
    ReDim Preserve argarray(NumArgs)
    GetCommandLine = argarray()
End Function

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

Public Function glPeriodYear(ByVal fy As Integer, ByVal fp As Byte, ByVal mon As Byte) As String
    Dim v As Variant
    If mon < fp Then
        v = DateSerial(fy - 1, mon, 1)
    Else
        v = DateSerial(fy, mon, 1)
    End If
    glPeriodYear = Format(v, "mmm yyyy")
End Function
