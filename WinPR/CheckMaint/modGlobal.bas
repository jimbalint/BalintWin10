Attribute VB_Name = "modGlobal"
Public PrvwReturn As Boolean
Public SQLString As String
Public cn As ADODB.Connection
Public cnPRCK As ADODB.Connection
Public Client As clsClient
Public Customer As clsCustomer
Public rsClient As New ADODB.Recordset
Public rsCustomer As New ADODB.Recordset
Public EditSw As Boolean
Public ClientID As Long
Public CustID As Long
Public ClientName As String

Public Ln As Integer
Public Lines As Integer
Public Columns As Integer
Public MaxLines As Integer
Public FontName As String
Public FontSize As Integer
Public Prvw As frmPreview
Public Progress As frmProgress
Private mPortrait As Byte
Private mLandscape As Byte

Public Condensed As Boolean
Public CharSize As Double
Dim CPI As Integer

Public XUnits As Integer
Public YUnits As Integer

Public PgNum As Long
Public FormatString(50) As String
Public PrintValue(50) As Variant
Public PS As String
Public FValue As Integer

Public X As Variant
Public TaskID As Long

Public DisConn As Boolean
Public SelID As Long

Public ModeSelect As Boolean

Public InitFlag As Boolean
Public PEDate, CheckDate, Startdate, EndDate As Date

Public RecAdd As Boolean
Public RecPut As Boolean
Public Portrait As Byte
Public HorzNudge, VertNudge, Nudge As Long

Public Sub SetEquates()
    RecAdd = True
    RecPut = False
    Equate.Portrait = 1
 
End Sub


Public Sub Prt(ByVal Line As Byte, ByVal Col As Byte, ByVal Str As String)

Dim pi As Integer

       Prvw.vsp.CurrentX = XUnits * Col + 200
       Prvw.vsp.CurrentY = YUnits * Line
       Prvw.vsp.Text = Str
       
'       Ln = Ln + 1

'       Printer.CurrentX = (Col * 120) - 240 + (hadj * 45)
'       Printer.CurrentY = ((Line - 1) * 240) - 960 + (vadj * 45)
    
End Sub

' col as actual twip value
Public Sub PrtCenter(ByVal Line As Byte, ByVal Str As String)

Dim pi As Integer
            
    I = ((Columns - Len(Str)) / 2)
    PrintValue(1) = " ":            FormatString(1) = "a" & I
    PrintValue(2) = Str:     FormatString(2) = "a" & Len(Str)
    PrintValue(3) = " ":            FormatString(3) = "~"
    FormatPrint
    
End Sub


Public Sub FormFeed()

   Prvw.vsp.NewPage
   Ln = 0

End Sub

Public Sub LineFeed(ByVal Lines As Integer)
    
Dim lx As Integer
    
    PrintValue(1) = " "
    FormatString(1) = "a1"
    PrintValue(2) = " "
    FormatString(2) = "~"

    For lx = 1 To Lines
        FormatPrint
    Next lx

End Sub


Public Sub PrtInit(ByVal PortLand As String)
   
   Set Prvw = New frmPreview
   
   Prvw.vsp.Preview = True
   
   Prvw.vsp.PhysicalPage = False
   
   Prvw.vsp.MarginRight = 0
   Prvw.vsp.MarginLeft = 0
   Prvw.vsp.MarginBottom = 0
   Prvw.vsp.MarginTop = 0
   
   If PortLand = "Port" Then
      Prvw.vsp.Orientation = orPortrait
   Else
      Prvw.vsp.Orientation = orLandscape
   End If
   
   Prvw.vsp.Font.Name = "Courier New"
      
   Prvw.Height = Screen.Height
   Prvw.Width = Prvw.vsp.Width + 1000
   Prvw.Top = 0
   Prvw.Left = 0

   Prvw.vsp.Height = Screen.Height - 500
   Prvw.vsp.Left = 500
   Prvw.vsp.Top = 0
 
   Prvw.vsp.StartDoc

   YUnits = 240 - 15       ' = 1440 / 6


End Sub
Public Sub SetFont(ByVal FSize As Integer, ByVal PortLand As Byte)
   
   Prvw.vsp.FontSize = FSize
   
   Select Case FSize
      Case 6
         XUnits = 72   ' 1440 / 20
         CPI = 20
      Case 7
         XUnits = 80   ' 1440 / 18
         CPI = 18
      Case 8
         XUnits = 96   ' 1440 / 15
         CPI = 15
      Case 9
         XUnits = 110  ' 1440 / 13
         CPI = 13
      Case 10
         XUnits = 120  ' 1440 / 12
         CPI = 12
      Case 11
         XUnits = 131  ' 1440 / 11
         CPI = 11
      Case 12
         XUnits = 144  ' 1440 / 10
         CPI = 10
      Case 13
         XUnits = 180   ' 1440 / 8   for Report of wages supplemental
         CPI = 8
      Case Else
         MsgBox "Incorrect font size selected: " & FSize
         End
   End Select
   
   ' store the width of one character in twips
   ' using a monospaced font - will be the same size no matter what the letter
   CharSize = Prvw.vsp.TextWidth("W")
   
   If PortLand = Portrait Then
      Prvw.vsp.Orientation = orPortrait
      MaxLines = 63
      Columns = 8 * CPI
   Else
      Prvw.vsp.Orientation = orLandscape
      MaxLines = 49
      Columns = 11 * CPI
   End If
   
End Sub


Public Sub GoBack()
   
    ' return call if given
    If BackName <> "" Then
        X = BackName & " UserID=" & UserID & " dbPwd=" & dbPwd & " OpenTab=" & OpenTab
        TaskID = Shell(X, vbMaximizedFocus)
    End If

    End

End Sub

Public Sub SetGrid(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    gfg.FixedCols = 0                   ' see all cols selected by SQL
    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    gfg.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set gfg.DataSource = grs.DataSource '
    gfg.DataMember = grs.DataMember     '

    gfg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    gfg.TabBehavior = flexTabCells                       ' tab moves between cells
    gfg.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    ' gfg.HighLight = flexHighlightNever                   ' don't select ranges

End Sub
Public Sub SetGridFree(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    gfg.FixedCols = 0                   ' see all cols selected by SQL
    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    gfg.DataMode = flexDMFree         ' Recordset cursor is maintained by grid
    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set gfg.DataSource = grs.DataSource '
    gfg.DataMember = grs.DataMember     '

    gfg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    gfg.TabBehavior = flexTabCells                       ' tab moves between cells
    gfg.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    ' gfg.HighLight = flexHighlightNever                   ' don't select ranges

End Sub

Public Sub AddAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    grs.AddNew          ' Add to the recordset
    grs.Update          ' Record (save to file)
    grs.MoveLast        ' Move to the last record in the record set
    
    gfg.DataRefresh     ' Update the grid data
    gfg.Col = 0         ' Go to the first column
    gfg.SetFocus        ' Move from add button to grid

End Sub
Public Sub DelAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
' Public Sub DelAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid, ByVal Number As Long)
    
    grs.Delete
    grs.Update          ' Record (save to file)
    grs.MoveLast        ' Move to the last record in the record set
    gfg.DataRefresh     ' Update the grid data
    gfg.Col = 0         ' Go to the first column
    gfg.SetFocus        ' Move from add button to grid
End Sub

Public Sub SetAdo(ByRef gcn As ADODB.Connection, ByRef grs As ADODB.Recordset, ByVal SQL As String)
    ' Common behavior for Recordsets
    Set grs = New ADODB.Recordset       ' set the recordset
    grs.LockType = adLockOptimistic     '
    grs.CursorType = adOpenDynamic      '
    grs.Source = SQL                    '
    Set grs.ActiveConnection = gcn      ' connection set previous to call
    grs.Open                            ' start the record
End Sub


Public Sub tdbTextSet(ByRef tdbTxt As TDBText)

    tdbTxt.Key.Clear = ""       ' no key to clear field
    tdbTxt.FormatMode = dbiIncludeFormat
    tdbTxt.Format = "A9#@"

End Sub

Public Sub tdbAmountSet(ByRef tdbAmt As TDBNumber)

    tdbAmt.Format = "##,###,##0.00;(##,###,##0.00)"
    tdbAmt.DisplayFormat = "##,###,##0.00;(##,###,##0.00);0.00"
    tdbAmt.HighlightText = True
    tdbAmt.Key.Clear = ""
    tdbAmt.MinValue = -99999999.99
    tdbAmt.MaxValue = 99999999.99
    tdbAmt.Value = 0

End Sub

Public Sub tdbIntegerSet(ByRef tdbAmt As TDBNumber)

    tdbAmt.Format = "##,###,##0;(##,###,##0)"
    tdbAmt.DisplayFormat = "##,###,##0;(##,###,##0);0"
    tdbAmt.HighlightText = True
    tdbAmt.Key.Clear = ""
    tdbAmt.MinValue = -99999999
    tdbAmt.MaxValue = 99999999

End Sub

Public Sub tdbDateSet(ByRef tdbDate As tdbDate, ByVal DateValue As Date)

    tdbDate.Format = "mm/dd/yyyy"
    tdbDate.DisplayFormat = "mm/dd/yyyy"
    tdbDate.ErrorBeep = True
    tdbDate.DropDown.Enabled = True
    tdbDate.DropDown.Visible = dbiShowAlways
    tdbDate.DropDown.Position = dbiDropPosInside

    If IsNull(DateValue) Then
        tdbDate.Text = ""
    ElseIf DateValue = 0 Then
        tdbDate.Text = ""
    Else
        tdbDate.Value = DateValue
    End If

End Sub

Public Function nNull(ByVal InVal As Variant) As Variant

    If IsNull(InVal) Then
        nNull = 0
    Else
        nNull = InVal
    End If

End Function

Public Function MonthName(ByVal MonthNum) As String

    Select Case MonthNum
        Case 1
            MonthName = "JANUARY"
        Case 2
            MonthName = "FEBRUARY"
        Case 3
            MonthName = "MARCH"
        Case 4
            MonthName = "APRIL"
        Case 5
            MonthName = "MAY"
        Case 6
            MonthName = "JUNE"
        Case 7
            MonthName = "JULY"
        Case 8
            MonthName = "AUGUST"
        Case 9
            MonthName = "SEPTEMBER"
        Case 10
            MonthName = "OCTOBER"
        Case 11
            MonthName = "NOVEMBER"
        Case 12
            MonthName = "DECEMBER"
        Case Else
            MonthName = ""
        End Select

End Function

Public Function GetNumMon(ByVal MonthAbbrev As String) As Integer

    MonthAbbrev = StrConv(MonthAbbrev, vbLowerCase)

    Select Case MonthAbbrev
        Case "jan"
            GetNumMon = 1
        Case "feb"
            GetNumMon = 2
        Case "mar"
            GetNumMon = 3
        Case "apr"
            GetNumMon = 4
        Case "may"
            GetNumMon = 5
        Case "jun"
            GetNumMon = 6
        Case "jul"
            GetNumMon = 7
        Case "aug"
            GetNumMon = 8
        Case "sep"
            GetNumMon = 9
        Case "oct"
            GetNumMon = 10
        Case "nov"
            GetNumMon = 11
        Case "dec"
            GetNumMon = 12
        Case Else
            GetNumMon = 0
    End Select

End Function

Public Function PadString(ByVal InString As String, ByVal StrLen As Integer, Optional ByVal Justify As String) As String

    ' left justified by default
    If Len(InString) > StrLen Then
        PadString = Mid(InString, 1, StrLen)
    Else
        If Justify = "R" Then
            PadString = Space(StrLen - Len(InString)) & InString
        Else
            PadString = InString & Space(StrLen - Len(InString))
        End If
    End If

End Function

Public Function OutNumber(ByVal InNumber As Long, ByVal StrLen As Integer) As String

Dim nString As String

    If InNumber < 0 Then
        MsgBox "OutNumber can not process a negative value! " & InNumber, vbCritical
        End
    End If

    nString = CStr(InNumber)
    
    If Len(nString) > StrLen Then
        MsgBox "OutNumber can not process an excess value! " & InNumber, vbCritical
        End
    End If

    OutNumber = String(StrLen - Len(CStr(InNumber)), "0") & CStr(InNumber)
    
End Function

Public Function ABACheckDigit(ByVal InString As String) As Byte

Dim CheckSum As Long
Dim Weight As Byte
    
    ABACheckDigit = 99

    If Len(InString) <> 9 Then
        MsgBox "Invalid length: " & Len(InString) & " " & InString, vbCritical
        Exit Function
    End If
    
    ' must be all numeric
    For I = 1 To 8
        
        If InStr("0123456789", Mid(InString, I, 1)) = 0 Then
            MsgBox "Invalid ABA - Numbers only! " & InString, vbCritical
            Exit Function
        End If
    
        If I = 1 Or I = 4 Or I = 7 Then
            Weight = 3
        ElseIf I = 2 Or I = 5 Or I = 8 Then
            Weight = 7
        Else
            Weight = 1
        End If
        
        CheckSum = CheckSum + Weight * CByte(Mid(InString, I, 1))
        
    Next I
    
    ABACheckDigit = CheckSum Mod 10
    If ABACheckDigit > 0 Then ABACheckDigit = 10 - ABACheckDigit

    ' compare to given value
    If ABACheckDigit <> CByte(Right(InString, 1)) Then
        MsgBox "Invalid Check Digit!", vbCritical
        ABACheckDigit = 99
        Exit Function
    End If

End Function

Public Sub GetNudge(ByVal UserID As Long, _
                    ByVal ReportName As String)
                    
    SQLString = "SELECT * FROM PRGlobal WHERE " & _
                "TypeCode = " & PREquate.GlobalTypeNudge & " AND " & _
                "Description = '" & ReportName & "' AND " & _
                "Year = " & UserID
    
    If PRGlobal.GetBySQL(SQLString) Then
        HorzNudge = PRGlobal.Flag
        VertNudge = PRGlobal.Month
    Else
        HorzNudge = 0
        VertNudge = 0
    End If
 
End Sub

Public Sub SaveNudge(ByVal UserID As Long, _
                     ByVal ReportName As String)
                         
    SQLString = "SELECT * FROM PRGlobal WHERE " & _
                "TypeCode = " & PREquate.GlobalTypeNudge & " AND " & _
                "Description = '" & ReportName & "' AND " & _
                "Year = " & UserID
    
    If PRGlobal.GetBySQL(SQLString) Then
        PRGlobal.Flag = HorzNudge
        PRGlobal.Month = VertNudge
        PRGlobal.Save (Equate.RecPut)
    Else
        If HorzNudge = 0 And VertNudge = 0 Then     ' don't create if value is zero
            Exit Sub
        Else
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeNudge
            PRGlobal.Description = ReportName
            PRGlobal.Flag = HorzNudge
            PRGlobal.Month = VertNudge
            PRGlobal.Year = User.ID
            PRGlobal.Save (Equate.RecAdd)
        End If
    End If

End Sub

Public Sub SetNudge(ByRef tdbNum As TDBNumber)
    tdbIntegerSet tdbNum
    With tdbNum
        .Spin = dbiShowAlways
        .MinValue = 0
        .MaxValue = 255
    End With
End Sub


Public Sub rsInit(ByVal SQLString As String, _
                  ByRef cni As ADODB.Connection, _
                  ByRef rsi As ADODB.Recordset)
   
    Set rsi = New ADODB.Recordset
    rsi.Source = SQLString

    rsi.ActiveConnection = cni
   
    ' open as disconnected - client side
    If DisConn Then
        rsi.CursorLocation = adUseClient
        rsi.CursorType = adOpenKeyset
        rsi.LockType = adLockBatchOptimistic
    Else
        rsi.CursorLocation = adUseServer
        rsi.CursorType = adOpenKeyset
        rsi.LockType = adLockOptimistic
    End If
   
    rsi.Open

    If DisConn Then
        Set rsi.ActiveConnection = Nothing
        DisConn = False
    End If

End Sub

Public Sub FormatPrint()

Dim pi As Integer
Dim Lngth As Integer
Dim PCol As Integer
Dim RightString As String
Dim l As Integer

Dim DWidth As Integer
Dim DExp As Integer
Dim CCount As Integer

Dim Quote As String
Dim Comma As String
Dim TextString As String

Dim TxtX As String

    ' init variables for csv output
    Quote = """"
    Comma = ","

   pi = 1
   PS = ""
   PCol = 1
   
   If IsNull(CCol) Then
      PrintString = ""
   ElseIf CCol = 0 Then
      PrintString = ""
   Else
      PrintString = Space(CCol)
   End If

   TextString = ""

   Do
                    
      Lngth = Len(FormatString(pi))
      
      FValue = CInt(Mid(FormatString(pi), 2, Lngth - 1))
      
      Select Case Mid(FormatString(pi), 1, 1)
         
         Case "x"
         
            PrintString = PrintString & Space(FValue)
         
         Case "t"
            
            Lngth = Len(PrintString)
            If FValue <= Lngth Then
               PrintString = Mid(PrintString, 1, FValue - 1)
            Else
               PrintString = PrintString & Space(FValue - Lngth - 1)
            End If
            
         Case "a"
  
            l = Len(PrintValue(pi))
            If l < FValue Then
               PrintValue(pi) = PrintValue(pi) & Space(FValue - l)
            End If
            PrintString = PrintString & Mid(PrintValue(pi), 1, FValue)
            
            TextString = TextString & Quote & Mid(PrintValue(pi), 1, FValue) & Quote & Comma
            
         Case "r"       ' right justified
         
            If FValue > Len(PrintValue(pi)) Then
               PrintString = PrintString & Space(FValue - Len(PrintValue(pi))) & PrintValue(pi)
               TextString = TextString & Quote & Space(FValue - Len(PrintValue(pi))) & PrintValue(pi) & Quote & Comma
            Else
               PrintString = PrintString & Mid(PrintValue(pi), 1, FValue)
               TextString = TextString & Quote & Mid(PrintValue(pi), 1, FValue) & Quote & Comma
            End If
         
         Case "d"
            
            DWidth = 5         ' 0.00-
            DExp = 1
            CCount = 0
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
               If DExp Mod 3 = 1 And DExp <> 1 Then CCount = CCount + 1 ' comma count
            Loop
            DWidth = DWidth + CCount       ' add in the comma count
               
            X = Format(Abs(PrintValue(pi)), "##,###,##0.00")
            If PrintValue(pi) < 0 Then
               X = X & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0.00")    ' leading minus sign for text output
            Else
               X = X & Space(1)
               TxtX = X
            End If
               
            If DWidth <= 14 Then
                PrintString = PrintString & Space(14 - DWidth) & X
                TextString = TextString & Quote & Space(14 - DWidth) & TxtX & Quote & Comma
            Else
                PrintString = PrintString & X
                TextString = TextString & Quote & TxtX & Quote & Comma
            End If
         
         Case "p"
               
            DWidth = 4  ' 0.0-
            DExp = 1
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
            Loop
            
            X = Format(Abs(PrintValue(pi)), "##0.0")

            If PrintValue(pi) < 0 Then
               X = X & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##0.0")
            Else
               X = X & Space(1)
               TxtX = X
            End If
            
            If DWidth <= 6 Then
               PrintString = PrintString & Space(6 - DWidth) & X
               TextString = TextString & Quote & Space(6 - DWidth) & TxtX & Quote & Comma
            Else
               PrintString = PrintString & X
               TextString = TextString & Quote & TxtX & Quote & Comma
            End If
            
            PCol = PCol + FValue
            
         Case "q"     ' same as p except blank if zero
               
            If PrintValue(pi) <> 0 Then
            
                X = Format(Abs(PrintValue(pi)), "##0.0")

                If PrintValue(pi) < 0 Then
                   X = X & "-"
                   TxtX = "-" & Format(Abs(PrintValue(pi)), "##0.0")
                Else
                   X = X & Space(1)
                   TxtX = X
                End If
            
                X = Format(Abs(PrintValue(pi)), "##0.0")
            
                DWidth = 4  ' 0.0-
                DExp = 1
                Do
                  If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
                  DExp = DExp + 1
                  DWidth = DWidth + 1
                Loop
            
                PrintString = PrintString & Space(6 - DWidth) & X
                
                TextString = TextString & Quote & Space(6 - DWidth) & TxtX & Quote & Comma
            
            Else
            
                PrintString = PrintString & Space(FValue)
                TextString = TextString & Quote & Quote & Comma
            
            End If
         
         Case "i"
               
            DWidth = 2         ' 0-
            DExp = 1
            CCount = 0
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
               If DExp Mod 3 = 1 And DExp <> 1 Then CCount = CCount + 1 ' comma count
            Loop
            DWidth = DWidth + CCount       ' add in the comma count
               
            X = Format(Abs(PrintValue(pi)), "##,###,##0")

            If PrintValue(pi) < 0 Then
               X = X & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0")
            Else
               X = X & Space(1)
               TxtX = X
            End If
            
            PrintString = PrintString & Space(14 - DWidth) & X
            
            TextString = TextString & Comma & Space(14 - DWidth) & TxtX & Quote & Comma
            
         Case "n"
               
            X = Format(Abs(PrintValue(pi)), "########0")
                        
            PrintString = PrintString & Space(FValue - Len(X)) & X
            TextString = TextString & Quote & Space(FValue - Len(X)) & X & Quote & Comma
            
         Case Else
            
            MsgBox "Bad Format: " & FormatString(pi)
      
      End Select
      
      pi = pi + 1
   
   Loop Until FormatString(pi) = "~"
       
'   ' blank out columns 1 & 3 if suppr current period flag is set
'   If SupprFlag Then
'      If Len(PrintString) < 57 Then
'         PrintString = PrintString & Space(132 - Len(PrintString))
'      End If
'      PrintString = Mid(PrintString, 1, 34) & Space(22) & _
'                    Mid(PrintString, 57, 22) & Space(22) & _
'                    Mid(PrintString, 103, 29)
'
'   End If
       
   ' print the string
   Prvw.vsp.CurrentY = Ln * YUnits
   Prvw.vsp.CurrentX = XUnits * 2
   Prvw.vsp.Text = PrintString
       
   ' clear the print values
   For pi = 1 To 30
       PrintValue(pi) = ""
   Next pi

   ' output to text if necessary
   If TextChannel <> 0 Then
      Print #TextChannel, TextString
   End If

End Sub

Public Sub PosPrint(ByVal CurrX As Long, ByVal CurrY As Long, ByVal PrintString As String)
    Prvw.vsp.CurrentX = CurrX + HorzNudge * Nudge
    Prvw.vsp.CurrentY = CurrY + VertNudge * Nudge
    Prvw.vsp.Text = PrintString
End Sub
