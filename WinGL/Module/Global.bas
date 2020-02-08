Attribute VB_Name = "Global"
Option Explicit

Public HAdj As Single

Public Equate As cEquate
Public GLAccount As cGLAccount
Public GLAmount As cGLAmount
Public GLBranch As cGLBranch
Public GLCompany As cGLCompany
Public GLDescription As cGLDescription
Public GLHistory As cGLHistory
Public GLJournal As cGLJournal
Public GLPrint As cGLPrint
Public GLUser As cGLUser
Public GLBatch As cGLBatch
Public GLFFSched As cGLFFSched
Public GLFFColumn As cGLFFColumn
' Public GLColumn As cGLColumn

Public PRCompany As cPRCompany
Public PREmployee As cPREmployee
Public PRFWTTable As cPRFWTTable
Public PRState As cPRState
Public PREquate As cPREquate
Public PRGlobal As cPRGlobal
Public PRHist As cPRHist
Public PRDist As cPRDist
Public PRCity As cPRCity

Public MsgResponse As Integer
Public Response As Boolean
Public CompFlag As Boolean

Public ErrMessage As String
Public Ln As Integer
Public Lines As Integer
Public Columns As Integer
Public MaxLines As Integer
Public FontName As String
Public FontSize As Integer
Public Prvw As frmPreview
Public Condensed As Boolean
Public CharSize As Double
Dim CPI As Integer

Public XUnits As Integer
Public YUnits As Integer

Public PgNum As Long
Public FormatString(30) As String
Public PrintValue(30) As Variant
Public PS As String
Public FValue As Integer

Public ii As Long
Public jj As Long
Public kk As Long
Public xx As String
Public x As Variant

Public CurrYrPdEnd As String              ' g[1]
Public CurrYrCurrPdBeg As String          ' g[2]
Public CurrYrFYBeg As String              ' g[3]
Public CurrYrFYEnd As String              '     -------
Public PrevYrPdEnd As String              ' g[4]
Public PrevYrPDBeg As String              ' g[5]
Public PrevYrFYBeg As String              ' g[6]
Public ReportDate As String               ' g[7]

' command line variables
Public CompanyID As Long
Public dbPwd As String
Public ProgName As String
Public SysFile As String
Public UserID As Long
Public BatchNum As Long
Public BackName As String
Public Period As Long

Public TaskID As Long

Public FormType As Byte        ' 1=add / 2=edit / 3=delete ....

Public uDB As XArrayDB

Public PrintString As String
Public CCol As Long

Public TextChannel As Integer
Public TextFileName As String

Public SupprFlag As Boolean
Public LastH As String

Public SQLString As String
Public DBName As String
Public NoFieldCheck As Boolean
Public OpenTab As Byte

Public PrvwReturn As Boolean
Public BalintFolder As String

Public TabUnit, TabValue As Integer

Public TestMode As Boolean
Public Password As String
Public MenuName As String

Public Sub GetDates(ByVal ID As Long)
   
      ' get company record
      GLCompany.GetData (ID)
      
      ' curr yr pd ending     1
      CurrYrPdEnd = DateSerial(Int(GLPrint.EndDate / 100), _
                               GLPrint.EndDate Mod 100, 1)
      CurrYrPdEnd = DateAdd("m", 1, CurrYrPdEnd)
      CurrYrPdEnd = DateAdd("d", -1, CurrYrPdEnd)
   
      ' curr yr pd beg        2
      CurrYrCurrPdBeg = DateSerial(Int(GLPrint.BeginDate / 100), _
                               GLPrint.BeginDate Mod 100, 1)
                               
      ' curr yr fy beg        3
      CurrYrFYBeg = DateSerial(Int(GLPrint.BeginDate / 100), _
                          GLCompany.FirstPeriod, 1)
      
      If GLCompany.FirstPeriod > GLPrint.BeginDate Mod 100 Then
         CurrYrFYBeg = DateAdd("yyyy", -1, CurrYrFYBeg)
      End If
      
      ' FY end date
      CurrYrFYEnd = DateAdd("yyyy", 1, CurrYrFYBeg)
      CurrYrFYEnd = DateAdd("d", -1, CurrYrFYEnd)
      
      ' prev yr pd end        4
      PrevYrPdEnd = DateAdd("yyyy", -1, CurrYrPdEnd)
      
      ' prev yr pd beg        5
      PrevYrPDBeg = DateAdd("yyyy", -1, CurrYrCurrPdBeg)
      
      ' prev yr fy beg        6
      PrevYrFYBeg = DateAdd("yyyy", -1, CurrYrFYBeg)
      
End Sub

Public Function GetPeriod(ByVal NumPer As Integer, _
                          ByVal CalMonth As Integer, _
                          ByVal FYBeg As Integer) As Integer
   
   GetPeriod = ((NumPer + CalMonth - FYBeg) Mod NumPer) + 1

End Function

Public Function Div0(ByVal Numerator As Currency, _
                           Denominator As Currency) As Currency
                           
   If Denominator <> 0 Then
      Div0 = Numerator * 100 / Denominator
   Else
      Div0 = 0
   End If
                           
End Function


Public Function Centered(ByVal Strg As String, _
                         ByVal FldWidth As Integer) As String
                         
    If (FldWidth - Len(Strg)) < 0 Then
       Centered = Strg
    Else
       Centered = Space((FldWidth - Len(Strg)) / 2) & Strg
    End If
    
End Function

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

       Prvw.vsp.CurrentX = ((Equate.PgTwips - Prvw.vsp.TextWidth(Str)) / 2) + 200 + (TabUnit * TabValue)
       Prvw.vsp.CurrentY = 240 * Line
       Prvw.vsp.Text = Str
       
'       Ln = Ln + 1

       ' clear the print values
       For pi = 1 To 30
           PrintValue(pi) = ""
       Next pi

'       Printer.CurrentX = (Col * 120) - 240 + (hadj * 45)
'       Printer.CurrentY = ((Line - 1) * 240) - 960 + (vadj * 45)
    
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
'   Prvw.vsp.PhysicalPage = True
   
   Prvw.vsp.MarginRight = 0
   Prvw.vsp.MarginLeft = 0
   Prvw.vsp.MarginBottom = 0
   Prvw.vsp.MarginTop = 0
   
   If PortLand = "Port" Then
      Prvw.vsp.Orientation = orPortrait
   Else
      Prvw.vsp.Orientation = orLandscape
   End If
   
' -------- from glstmt
'   If GLPrint.PrintBIB = Equate.PrtISOnly And GLPrint.WidePrint = False Then
'      Prvw.vsp.Orientation = orLandscape
'   End If
' --------
   
   Prvw.vsp.Font.Name = "Courier New"
   
   ' bold for Richlak - make an option ....
   If GLPrint.Output = "Bold" Then
      Prvw.vsp.Font.Bold = True
   Else
      Prvw.vsp.Font.Bold = False
   End If
    
   Prvw.Height = Screen.Height
   Prvw.Width = Prvw.vsp.Width + 1000
   Prvw.Top = 0
   Prvw.Left = 0

   Prvw.vsp.Height = Screen.Height - 500
   Prvw.vsp.Left = 500
   Prvw.vsp.Top = 0
 
'   Prvw.vsp.ExportFormat = vpxPagedHTML
'   Prvw.vsp.ExportFile = "c:\asend\htmltest.html"
 
   Prvw.vsp.StartDoc

   YUnits = 240 - 15       ' = 1440 / 6

'   YUnits = 240 - 25

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
      Case Else
         MsgBox "Incorrect font size selected: " & FSize
         End
   End Select
   
   ' store the width of one character in twips
   ' using a monospaced font - will be the same size no matter what the letter
   CharSize = Prvw.vsp.TextWidth("W")
   
   If PortLand = Equate.Portrait Then
      Prvw.vsp.Orientation = orPortrait
      MaxLines = 63
      Columns = 8 * CPI
   Else
      Prvw.vsp.Orientation = orLandscape
      MaxLines = 49
      Columns = 11 * CPI
   End If
   
End Sub


Public Sub SetEquates()
   
    TabUnit = 45    ' 1/32 "
    
    Equate.Regular = 1
    Equate.Branch = 2
    Equate.Consol = 3
    Equate.Budget = 4
   
    Equate.NonComp = 1
    Equate.Comp = 2
   
    Equate.Stmt = 1
    Equate.Sched = 2
   
    Equate.PrtBSOnly = 1
    Equate.PrtISOnly = 2
    Equate.PrtBoth = 3

    Equate.PgTwips = 11520 ' 8" x 1440 twips/inch = 11520

    Equate.RecAdd = True
    Equate.RecPut = False

    Equate.FormAdd = 1
    Equate.FormEdit = 2
    Equate.FormDel = 3

    Equate.Portrait = 1
    Equate.LandScape = 2

    Equate.OpAdd = 0
    Equate.OpSub = 1
    Equate.OpPct = 2
    Equate.OpMult = 3
    Equate.OpDiv = 4

    Equate.ColAvg = 1
    Equate.ColProj = 2
    Equate.ColMultiply = 3
    Equate.ColAdd = 4
    Equate.ColSubtract = 5
    Equate.ColDivide = 6
    Equate.ColCurrPd = 7
    Equate.ColPriorPd = 8
    Equate.ColYTD = 9
    Equate.ColAllPd = 10
    Equate.ColCustom = 11

    ' ------------ added 2016-01-02 ------------
    ' --- needed for tax updates
    PREquate.GlobalTypeRaceCode = 1
    PREquate.GlobalTypeEducationLevel = 2
    PREquate.GlobalTypeContact = 3
    PREquate.GlobalTypeShiftCode = 4
    PREquate.GlobalTypeTerminationCode = 5
    
    PREquate.GlobalTypeSSMax = 6
    PREquate.GlobalTypeSSPct = 7
    PREquate.GlobalTypeMEDPct = 8
    PREquate.GlobalTypeFWTAllow = 9
    PREquate.GLobalTypeOHAllow = 10
    PREquate.GlobalTypeFUNMax = 11
    ' ------------ added 2016-01-02 ------------
    
    PREquate.GlobalTypeCompanyOption = 55
    PREquate.GlobalTypeGLFFSched = 59
    PREquate.GlobalTypeGLFFColumn = 61
    PREquate.GlobalTypeGLFFSetup = 62

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
               
            x = Format(Abs(PrintValue(pi)), "##,###,##0.00")
            If PrintValue(pi) < 0 Then
               x = x & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0.00")    ' leading minus sign for text output
            Else
               x = x & Space(1)
               TxtX = x
            End If
               
            If DWidth <= 14 Then
                PrintString = PrintString & Space(14 - DWidth) & x
                TextString = TextString & Quote & Space(14 - DWidth) & TxtX & Quote & Comma
            Else
                PrintString = PrintString & x
                TextString = TextString & Quote & TxtX & Quote & Comma
            End If
         
         Case "p"
               
            PrintValue(pi) = Round(PrintValue(pi), 1)
               
            DWidth = 4  ' 0.0-
            DExp = 1
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
            Loop
            
            x = Format(Abs(PrintValue(pi)), "##0.0")

            If PrintValue(pi) < 0 Then
               x = x & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##0.0")
            Else
               x = x & Space(1)
               TxtX = x
            End If
            
            If DWidth <= 6 Then
               PrintString = PrintString & Space(6 - DWidth) & x
               TextString = TextString & Quote & Space(6 - DWidth) & TxtX & Quote & Comma
            Else
               PrintString = PrintString & x
               TextString = TextString & Quote & TxtX & Quote & Comma
            End If
            
            PCol = PCol + FValue
            
         Case "q"     ' same as p except blank if zero
               
            If PrintValue(pi) <> 0 Then
            
                x = Format(Abs(PrintValue(pi)), "##0.0")

                If PrintValue(pi) < 0 Then
                   x = x & "-"
                   TxtX = "-" & Format(Abs(PrintValue(pi)), "##0.0")
                Else
                   x = x & Space(1)
                   TxtX = x
                End If
            
                x = Format(Abs(PrintValue(pi)), "##0.0")
            
                DWidth = 4  ' 0.0-
                DExp = 1
                Do
                  If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
                  DExp = DExp + 1
                  DWidth = DWidth + 1
                Loop
            
                PrintString = PrintString & Space(6 - DWidth) & x
                
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
               
            x = Format(Abs(PrintValue(pi)), "##,###,##0")

            If PrintValue(pi) < 0 Then
               x = x & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0")
            Else
               x = x & Space(1)
               TxtX = x
            End If
            
            PrintString = PrintString & Space(14 - DWidth) & x
            
            TextString = TextString & Comma & Space(14 - DWidth) & TxtX & Quote & Comma
            
         Case "n"
               
            x = Format(Abs(PrintValue(pi)), "########0")
                        
            PrintString = PrintString & Space(FValue - Len(x)) & x
            TextString = TextString & Quote & Space(FValue - Len(x)) & x & Quote & Comma
            
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
   
    ' include tab parameter
   Prvw.vsp.CurrentX = XUnits * 2 + (TabUnit * TabValue)
   
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

Public Function cmdline(ByVal x As String, _
                        ByRef ID As Long, _
                        ByRef Password As String, _
                        ByRef Prog As String, _
                        ByRef SysFile As String, _
                        ByRef User As String, _
                        ByRef BatchNum As Long) As Boolean
                        
' parse the command line in the following syntax:
'     id/password/progname/sysfile/user{/batch#}
                        
Dim ItemCount As Byte
Dim Posn As Integer
Dim I As Long
Dim Y As String
                        
    ID = 0
    Prog = ""
    Password = ""
    SysFile = ""
    User = ""
    BatchNum = 0
                        
    ItemCount = 1
    I = 0
    Do
        Posn = Posn + 1
        If Posn > Len(x) Then Exit Do
        If Mid(x, Posn, 1) = "/" Then
           If ItemCount = 1 Then ID = Y
           If ItemCount = 2 Then Password = Y
           If ItemCount = 3 Then Prog = Y
           If ItemCount = 4 Then SysFile = Y
           If ItemCount = 5 Then User = Y
           ItemCount = ItemCount + 1
           If Mid(x, Posn + 1, 1) <> "/" Then
              Posn = Posn + 1
           End If
           Y = ""
        End If
        Y = Y & Mid(x, Posn, 1)
    Loop
    
    If ItemCount = 5 Then User = Y
    If ItemCount = 6 Then BatchNum = Y
                            
    cmdline = True
    If ID = 0 Then cmdline = False
    If Prog = "" Then cmdline = False
    If SysFile = "" Then cmdline = False
    If User = "" Then cmdline = False
                        
End Function

Public Function GetCmd(ByVal cmdline As String, ByVal Argument As String, ByVal StrNum As String) As Variant

' return xxxx - Argument=xxxx

Dim I As Long
Dim cmd As String

    StrNum = LCase(StrNum)
    If StrNum <> "str" And StrNum <> "num" Then
       MsgBox "StrNum not assigned !"
       GetCmd = ""
       Exit Function
    End If
    
    If StrNum = "str" Then
       GetCmd = ""
    Else
       GetCmd = 0
    End If

    ' bad value traps
    If IsNull(cmdline) Then Exit Function
    If IsNull(Argument) Then Exit Function
    If cmdline = "" Then Exit Function
    If Argument = "" Then Exit Function

    ' ignore case for argument type but keep it for the return string
    cmd = LCase(cmdline)
    Argument = LCase(Argument)

    ' search for Argument=xxxxx
    I = InStr(1, cmd, Argument, vbTextCompare)
    If I = 0 Then Exit Function
    
    ' now look for the "=" sign
    If Mid(cmdline, I + Len(Argument), 1) <> "=" Then Exit Function
    
    ' append to make return string until a space or end of line
    I = I + Len(Argument) + 1
    Do
       If I > Len(cmdline) Then Exit Do
       If Mid(cmdline, I, 1) = " " Then Exit Do
       GetCmd = GetCmd & Mid(cmdline, I, 1)
       I = I + 1
    Loop

End Function


Public Sub Delay(ByVal Secs As Integer)

Dim n As Variant
Dim s1 As Long
Dim s2 As Long

    ' delay a specified number of seconds - what if spans a day?
    n = Now()
    s1 = (n - Int(n)) * 60 * 60 * 24   ' number of seconds since midnight
    
    Do
       n = Now()
       s2 = (n - Int(n)) * 60 * 60 * 24   ' number of seconds since midnight
       If s2 - s1 >= Secs Then Exit Do
    Loop
 
End Sub

Public Sub GoBack()
   
'    ' return call if given
'    If BackName <> "" Then
'
''        x = BackName & " UserID=" & UserID & _
''            " dbPwd=" & dbPwd & _
''            " OpenTab=" & OpenTab & _
''            " MenuName=" & MenuName
'
'        x = BackName & " UserID=" & UserID & _
'            " dbPwd=" & dbPwd & _
'            " OpenTab=" & OpenTab
'
'        If BalintFolder <> "" Then
'            x = x & " BalintFolder=" & BalintFolder
'        End If
'        TaskID = Shell(x, vbMaximizedFocus)
    
    ' return call if given
    If BackName <> "" Then
                
        x = BackName & " UserID=" & UserID & _
            " dbPwd=" & dbPwd & _
            " OpenTab=" & OpenTab & _
            " MenuName=" & MenuName
            
        If BalintFolder <> "" Then
            x = x & " BalintFolder=" & BalintFolder
        End If

        TaskID = Shell(x, vbMaximizedFocus)
    
    End If

    End

End Sub

Public Sub gGoBack()
   
   ' return call if given
   If BackName <> "" Then
      x = BackName & " UserID=" & UserID & " OpenTab=" & OpenTab
      If BalintFolder <> "" Then
          x = x & " BalintFolder=" & BalintFolder
      End If
      If dbPwd <> "" Then
          x = Trim(x) & " dbPwd=" & dbPwd
      End If
      
      TaskID = Shell(x, vbMaximizedFocus)
   End If

   End

End Sub


Public Function PeriodName(ByVal FY As Integer, _
                           ByVal Pd As Byte, _
                           ByVal FirstPD As Byte, _
                           ByVal NumPds As Byte) As String
                             
Dim v As Variant
                             
    If NumPds = 13 Then
       PeriodName = FY & " Period # " & Pd
       Exit Function
    End If
                             
    If FirstPD = 1 Then
       v = DateSerial(FY, Pd, 1)
    Else
       If Pd > 13 - FirstPD Then
          v = DateSerial(FY, Pd - 13 + FirstPD, 1)
       Else
          v = DateSerial(FY - 1, Pd + FirstPD - 1, 1)
       End If
    End If
    
    PeriodName = "Pd. #: " & Pd & " - " & Format(v, "mmmm-yyyy")
    
End Function

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

Public Function nNull(ByVal InVal As Variant) As Variant

    If IsNull(InVal) Then
        nNull = 0
    ElseIf InVal = "" Then
        nNull = 0
    Else
        nNull = InVal
    End If

End Function

Public Sub rsSave(rsDIS As ADODB.Recordset, cnn As ADODB.Connection)

    ' update a disconnected record set
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = cnn
    
    ' reopen the disconnected record set
    rs.Open rsDIS, cnn
    
    rs.UpdateBatch

End Sub

Public Function TableExists(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim FString As String
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
           
    TableExists = False
           
    Do Until frs.EOF = True
                  
        If frs!Table_Name = TableName Then
            TableExists = True
            Exit Do
        End If
        
       frs.MoveNext
   
   Loop

End Function


Public Sub cmbPoint(ByRef cmb As ComboBox, ByVal ID As Long)

Dim li As Long

    With cmb
        .Enabled = True
        If .ListCount = 0 Then Exit Sub
        .ListIndex = 0
        For li = 0 To .ListCount - 1
            If .ItemData(li) = ID Then
                .ListIndex = li
                Exit For
            End If
        Next li
    End With

End Sub

Public Function mdbName(ByVal Str As String) As String

Dim mdbI, mdbJ, mdbK As Long

    mdbName = ""
    If Str = "" Then Exit Function
    If InStr(1, Str, "\", vbTextCompare) = 0 Then Exit Function
    
    mdbK = Len(Str)
    For mdbI = mdbK To 1 Step -1
        If Mid(Str, mdbI, 1) = "\" Then
            Exit For
        End If
    Next mdbI
    If mdbI = 0 Then Exit Function
    mdbName = Trim(Mid(Str, mdbI + 1, mdbK))

End Function

