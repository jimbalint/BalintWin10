Attribute VB_Name = "modInvGlobal"
Option Explicit

Public InvStock As cInvStock
Public InvHeader As cInvHeader
Public InvBody As cInvBody
Public InvEquate As cInvEquate
Public InvGlobal As cInvGlobal

Public boo As Boolean

Public UseSalesTax As Boolean

Public VertAdj As Integer

' ------------------------------------------------------------------------------
' *** Direct Print Definitions ***

' ********************************************************************************************
' ********************************************************************************************
' ********************************************************************************************
' *** Advanced Printer Settings - PRINT DIRECT TO PRINTER - NO PRINT SPOOLER ***
' ********************************************************************************************
' ********************************************************************************************
' ********************************************************************************************

Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias _
   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
   pDocInfo As DOCINFO) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
   pcWritten As Long) As Long

Dim lhPrinter As Long
Dim lReturn As Long
Dim lpcWritten As Long
Dim lDoc As Long
Dim sWrittenData As String
Dim MyDocInfo As DOCINFO

Dim PrintAsPacking As Integer

' ------------------------------------------------------------------------------

' ******************************************************************************
' *** variables for KP invoice printing
Dim SoldShip(2, 5) As String
Dim PrinterName As String
Dim HAdj, VAdj As Byte
Dim PageNum As Long
Dim Ln As Long
Dim Dbl1, Dbl2, Dbl3 As Double

' ******************************************************************************

Dim i, j, k, l As Long
Dim x, y, z As String

Public Sub InvSetEquates()

    InvEquate.GlobalTypeTruck = 1
    InvEquate.GlobalTypeTrailer = 2
    InvEquate.GlobalTypeDriver = 3
    InvEquate.GlobalTypeTerms = 4
    InvEquate.GlobalTypeComment = 5
    InvEquate.GlobalTypeInvNumber = 6
    InvEquate.GlobalTypeQBSetup = 7
    InvEquate.GlobalTypeInvPrinter = 8
    InvEquate.GlobalTypeInvMessage = 9
    InvEquate.GlobalTypeSalesTax = 10
    InvEquate.GlobalTypeVAdj = 11
    InvEquate.GlobalTypeInvPrinter2 = 12
    
    InvEquate.IBMCPI10 = Chr(18)
    InvEquate.IBMCPI12 = Chr(27) & Chr(58)
    InvEquate.IBMCPI17 = Chr(18) & Chr(27) & Chr(15)
    InvEquate.IBMDblWide = Chr(27) & Chr(14)
    
End Sub

Public Function GetSecs(ByVal dte As Date) As Double

    ' number of seconds past midnite
    GetSecs = Round((dte - Int(dte)) * 86400, 0)

End Function

Public Function NumValue(ByVal Str As String) As Double

    NumValue = 0
    If IsNull(Str) Then Exit Function
    If Str = "" Then Exit Function
    If IsNumeric(Str) = False Then Exit Function
    NumValue = CDbl(Str)

End Function

Public Function KP_PrintInvoice(ByVal InvNum As Long, ByVal PrinterName As String, ByVal PrtAsPckg As Integer) As Boolean

Dim LastRow As Long
Dim TotalQO, TotalQS As Long
Dim InvTotal As Currency
Dim SubFlag As Boolean
Dim LPP As Byte

    PrintAsPacking = PrtAsPckg

    If DP_Init(PrinterName) = False Then Exit Function
    
    SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber = " & InvNum
    If InvHeader.GetBySQL(SQLString) = False Then
        MsgBox "Invoice Number Not Found: " & InvNum, vbExclamation
        KP_PrintInvoice = False
        Exit Function
    End If
    
    ' lines per page
    LPP = 64
    
    KP_PrintInvoice = True
    
    PageNum = 0
    
    KP_PrintHeader
    
    SQLString = " SELECT * FROM InvBody WHERE HeaderID = " & InvHeader.HeaderID & _
                " ORDER BY LineNum"
    
    ' no body info ???
    If InvBody.GetBySQL(SQLString) = False Then
        DP_PrintLine vbFormFeed
        DP_EndDoc
        Exit Function
    End If
    
    Do
    
        ' qty ordered
        If InvBody.QtyOrdered <> 0 Then
            x = StringPad(CStr(InvBody.QtyOrdered), 6, True)
        Else
            x = String(6, " ")
        End If
        
        x = x & String(2, " ")
        
        ' qty Shipped
        If InvBody.QtyShipped <> 0 Then
            x = x & StringPad(CStr(InvBody.QtyShipped), 6, True)
        Else
            x = x & String(6, " ")
        End If
        
        ' update the totals
        TotalQO = TotalQO + InvBody.QtyOrdered
        TotalQS = TotalQS + InvBody.QtyShipped
        InvTotal = InvTotal + InvBody.Amount
        
        ' description
        ' 2012-09-01 - strip control characters
        x = x & String(3, " ") & _
            StringPad(StripCtrlChar(InvBody.Description), 40)
        
        ' prices not on delivery slip
        If InvHeader.InvoiceDate <> 0 And PrintAsPacking = vbNo Then
        
            j = InvBody.QtyOrdered + InvBody.QtyShipped + InvBody.Price + InvBody.Amount
            If j <> 0 Then
            
                ' unit price
                x = x & String(3, " ")
                
                ' >>>> plain cents or hundreths of a cent
                Dbl1 = InvBody.Price
                If Dbl1 * 10 ^ 4 Mod 100 = 0 Then
                    y = Format(Dbl1, "###0.00")
                    x = x & StringPad(y, 9, True)
                Else
                    y = Format(Dbl1, "###0.0000")
                    x = x & StringPad(y, 9, True)
                End If
                
                ' amount
                Dbl1 = InvBody.Amount
                y = Format(Dbl1, "####,##0.00")
                x = x & StringPad(y, 11, True)
                        
            End If
        
        End If
        
        DP_PrintLine x
    
        Ln = Ln + 1
        If Ln > LPP - 1 Then KP_NextPage
            
        If InvBody.GetNext = False Then Exit Do
    
    Loop
    
    ' print appt date / time
    If InvHeader.TruckID1 <> 0 And InvHeader.ApptDate <> 0 Then
        
        If Ln > LPP - 3 Then KP_NextPage
        
        DP_LF
        Ln = Ln + 1
        
        x = String(17, " ") & "APPOINTMENT SCHEDULED FOR:"
        DP_PrintLine x
        
        ' X = String(17, " ") & Format(InvHeader.ApptDateTime, "h:mm AM/PM dddd mm/dd/yyyy")
        ' DP_PrintLine X
        
        x = String(17, " ") & InvHeader.ApptTime & " "
        x = x & Format(InvHeader.ApptDate, "dddd mm/dd/yyyy")
        DP_PrintLine x
    
        Ln = Ln + 2
    
    End If
    
    ' customer messages
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeInvMessage & _
                " AND CompanyID = " & PRCompany.CompanyID & _
                " AND UserID = " & InvHeader.SoldJobID
    If InvGlobal.GetBySQL(SQLString) = True Then
        j = 0
        For i = 1 To 5
            If i = 1 Then x = InvGlobal.Var1
            If i = 2 Then x = InvGlobal.Var2
            If i = 3 Then x = InvGlobal.Var3
            If i = 4 Then x = InvGlobal.Var4
            If i = 5 Then x = InvGlobal.Var5
            If x <> "" Then
                j = j + 1
                If j = 1 Then
                    DP_LF
                    Ln = Ln + 1
                End If
                y = String(17, " ") & x
                DP_PrintLine y
                Ln = Ln + 1
                If Ln > LPP - 1 Then KP_NextPage
            End If
        Next i
    End If
    
    If (InvHeader.InvoiceDate <> 0 And PrintAsPacking = vbNo) And InvHeader.SalesTax <> 0 Then
    
        If Ln > LPP - 3 Then KP_NextPage
        
        x = String(70, " ") & "----------"
        DP_PrintLine x
        Ln = Ln + 1
        
        x = String(58, " ") & " SUBTOTAL:" & String(1, " ")
        y = Format(InvTotal, "####,##0.00")
        x = x & StringPad(y, 11, True)
        DP_PrintLine x
        Ln = Ln + 1
        
        x = String(58, " ") & "SALES TAX:" & String(1, " ")
        y = Format(InvHeader.SalesTax, "####,##0.00")
        x = x & StringPad(y, 11, True)
        DP_PrintLine x
        Ln = Ln + 1
        
        InvTotal = InvTotal + InvHeader.SalesTax
    
    End If
    
    If (InvHeader.InvoiceDate <> 0 And PrintAsPacking = vbNo) And InvHeader.Freight <> 0 Then
        
        If Ln > LPP - 3 Then KP_NextPage
        
        If InvHeader.SalesTax = 0 Then
            x = String(70, " ") & "----------"
            DP_PrintLine x
            Ln = Ln + 1
            
            x = String(58, " ") & " SUBTOTAL:" & String(1, " ")
            y = Format(InvTotal, "####,##0.00")
            x = x & StringPad(y, 11, True)
            DP_PrintLine x
            Ln = Ln + 1
        End If
    
        x = String(58, " ") & " FREIGHT :" & String(1, " ")
        y = Format(InvHeader.Freight, "####,##0.00")
        x = x & StringPad(y, 11, True)
        DP_PrintLine x
        Ln = Ln + 1
        
        InvTotal = InvTotal + InvHeader.Freight
    
    End If
    
    If Ln > LPP - 1 Then KP_NextPage
    x = "------  ------"
    If (InvHeader.InvoiceDate <> 0 And PrintAsPacking = vbNo) Then
        x = x & String(56, " ") & "----------"
    End If
    DP_PrintLine x
    Ln = Ln + 1
    
    If TotalQO <> 0 Then
        x = StringPad(CStr(TotalQO), 6, True)
    Else
        x = String(6, " ")
    End If
    
    x = x & String(2, " ")
    
    ' qty Shipped
    If TotalQS <> 0 Then
        x = x & StringPad(CStr(TotalQS), 6, True)
    Else
        x = x & String(6, " ")
    End If
   
    x = x & String(3, " ") & "<====== TOTAL QUANTITY"
    
    If (InvHeader.InvoiceDate <> 0 And PrintAsPacking = vbNo) Then
        x = x & String(5, " ") & "INVOICE TOTAL ======>    "
        y = Format(InvTotal, "####,##0.00")
        x = x & StringPad(y, 11, True)
    End If
    
    If Ln > LPP - 1 Then KP_NextPage
    DP_PrintLine x
    
    DP_LF 2
    
    Ln = Ln + 3
    
    If Ln > LPP - 2 Then KP_NextPage
    
    x = String(17, " ") & "RECEIVED BY: ___________________________"
    DP_PrintLine x
    Ln = Ln + 1
    
    DP_LF
    Ln = Ln + 1
    
    If Ln > LPP - 3 Then KP_NextPage
    
    x = String(17, " ") & "TOTAL NUMBER OF PACKAGES: " & InvHeader.PackageCount
    DP_PrintLine x
    Ln = Ln + 1
    
    DP_LF 1
    Ln = Ln + 1
    
    x = String(17, " ") & "TOTAL NUMBER OF PALLETS:  " & InvHeader.PalletCount
    DP_PrintLine x
    
    DP_PrintLine vbFormFeed
    DP_EndDoc

'    PrtInit ("Port")
'    SetFont 10, Equate.Portrait
'
'    For i = 1 To 5
'        Ln = Ln + 1
'        PrintValue(1) = SoldShip(1, i):         FormatString(1) = "a50"
'        PrintValue(2) = SoldShip(2, i):         FormatString(2) = "a40"
'        PrintValue(3) = " ":                    FormatString(3) = "~"
'        FormatPrint
'    Next i
'
'    Prvw.vsp.EndDoc
'    Prvw.Show

End Function

Private Sub KP_NextPage()
            
    x = String(17, " ")
    If InvHeader.InvoiceDate = 0 Or PrintAsPacking = vbYes Then
        x = x & "***** DELIVERY SLIP CONTINUED ON NEXT PAGE *****"
    Else
        x = x & "***** INVOICE CONTINUED ON NEXT PAGE *****"
    End If
    DP_PrintLine x
    KP_PrintHeader
            
End Sub

Private Sub KP_PrintHeader()

Dim SoldCount, ShipCount As Byte
Dim SoldString(5) As String
Dim ShipString(5) As String
Dim VertAdjust As Integer

    If PageNum <> 0 Then DP_PrintLine vbFormFeed
    
    PageNum = PageNum + 1
    
    
    'DP_LF VAdj
    
    VertAdj = 0
    
    If VertAdj = 0 Then
        DP_LF       ' printer adjustment not entered
    Else
        
        x = Chr(27) & Chr(65) & "6" & Chr(27) & Chr(50)   ' set 6/72" = 1/12"
        DP_PrintLine x, True        ' print w/ no CR
        
        x = ""
        For i = 1 To VertAdj
            DP_PrintLine x
        Next i
        
        x = Chr(27) & Chr(65) & "12" & Chr(27) & Chr(50)   ' set 12/72" = 1/6"
        DP_PrintLine x, True
    
    End If
    
    x = InvEquate.IBMCPI10 & String(67, " ") & "PAGE: " & PageNum
    DP_PrintLine x
    
    DP_LF
    
    ' dbl wide - single line
    x = InvEquate.IBMDblWide & String(24, " ")
    If InvHeader.InvoiceDate = 0 Or PrintAsPacking = vbYes Then
        x = x & "DELIVERY SLIP"
    Else
        x = x & "      INVOICE"
    End If
    DP_PrintLine x
    
    DP_LF 2
    
    ' Inv Number
    x = InvEquate.IBMDblWide & String(31, " ") & InvHeader.InvoiceNumber
    DP_PrintLine x
    
    ' Date
    x = InvEquate.IBMDblWide & String(28, " ")
    If InvHeader.InvoiceDate = 0 Or PrintAsPacking = vbYes Then
        x = x & Format(InvHeader.OrderDate, "mm/dd/yyyy")
    Else
        x = x & Format(InvHeader.InvoiceDate, "mm/dd/yyyy")
    End If
    DP_PrintLine x
    
    DP_LF 3
    
    ' Sold / Ship To
    SoldCount = 0
    ShipCount = 0
    For i = 1 To 5
    
        ' sold to
        Select Case i
            Case 1:   x = InvHeader.SoldAddr1: y = InvHeader.ShipAddr1
            Case 2:   x = InvHeader.SoldAddr2: y = InvHeader.ShipAddr2
            Case 3:   x = InvHeader.SoldAddr3: y = InvHeader.ShipAddr3
            Case 4:   x = InvHeader.SoldAddr4: y = InvHeader.ShipAddr4
            Case 5
                
                x = InvHeader.SoldCity & InvHeader.SoldState & InvHeader.SoldZip
                If x <> "" Then
                    x = Trim(InvHeader.SoldCity) & ", " & Trim(InvHeader.SoldState) & "  " & InvHeader.SoldZip
                End If
                
                y = InvHeader.ShipCity & InvHeader.ShipState & InvHeader.ShipZip
                If y <> "" Then
                    y = Trim(InvHeader.ShipCity) & ", " & Trim(InvHeader.ShipState) & "  " & InvHeader.ShipZip
                End If
                        
        End Select
        
        If Trim(x) <> "" Then
            SoldCount = SoldCount + 1
            SoldString(SoldCount) = x
        End If
        
        If Trim(y) <> "" Then
            ShipCount = ShipCount + 1
            ShipString(ShipCount) = y
        End If
        
    Next i
    
    For i = 1 To 5
        x = String(9, " ") & StringPad(SoldString(i), 30) & _
            String(9, " ") & StringPad(ShipString(i), 30)
        DP_PrintLine x
    Next i
    
    DP_LF 2
    
    SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & InvEquate.GlobalTypeTerms & _
                " AND Var1 = '" & InvHeader.Terms & "'"
    If InvGlobal.GetBySQL(SQLString) = True Then
        y = InvGlobal.Description
    Else
        y = ""
    End If
    
    x = String(5, " ") & _
        StringPad(InvHeader.PO1, 20) & _
        String(5, " ") & _
        StringPad(InvHeader.PO2, 20) & _
        String(10, " ") & _
        y
    DP_PrintLine x
    
    DP_LF 3
    
    ' transportation box
    For i = 1 To 3
        DP_PrintLine KP_Transpo(i)
    Next i
        
    DP_LF 3
    
    Ln = 29
    
End Sub

Private Function KP_Transpo(ByVal num As Byte) As String
        
    KP_Transpo = ""
    
    If i = 1 Then
        j = InvHeader.TruckID1
        k = InvHeader.TrailerID1
        l = InvHeader.DriverID1
    End If
    
    If i = 2 Then
        j = InvHeader.TruckID2
        k = InvHeader.TrailerID2
        l = InvHeader.DriverID2
    End If
    
    If i = 3 Then
        j = InvHeader.TruckID3
        k = InvHeader.TrailerID3
        l = InvHeader.DriverID3
    End If
    
    If InvGlobal.GetByID(j) = False Then
        KP_Transpo = Space(27)
    Else
        KP_Transpo = StringPad(InvGlobal.Description, 23)
        KP_Transpo = KP_Transpo & String(4, " ")
    End If
    
    If InvGlobal.GetByID(k) = False Then
        KP_Transpo = KP_Transpo & Space(24)
    Else
        KP_Transpo = KP_Transpo & StringPad(InvGlobal.Description, 20)
        KP_Transpo = KP_Transpo & String(4, " ")
    End If
    
    If InvGlobal.GetByID(l) = False Then
        KP_Transpo = KP_Transpo & Space(25)
    Else
        KP_Transpo = KP_Transpo & StringPad(InvGlobal.Description, 25)
    End If
    
End Function

Public Function DP_Init(ByVal PrinterName As String) As Boolean

    lReturn = OpenPrinter(PrinterName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "Printer not found: " & PrinterName, vbExclamation
        DP_Init = False
        Exit Function
    End If
    
    MyDocInfo.pDocName = "AAAAAA"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)

    DP_Init = True

End Function

Public Sub DP_EndDoc()
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
End Sub

Public Sub DP_LF(Optional LineCount As Byte)

Dim Lines As Long
    
    If LineCount = 0 Then LineCount = 1
    For Lines = 1 To LineCount
        DP_PrintLine " "
    Next Lines

End Sub
Public Sub DP_PrintLine(ByVal Str As String, Optional SkipCR As Boolean)
    If SkipCR = False Then
        Str = Str & vbCrLf
    End If
    lReturn = WritePrinter(lhPrinter, ByVal Str, _
              Len(Str), lpcWritten)
End Sub
Private Function StringPad(ByVal Str As String, _
                           ByVal StrLen As Long, _
                           Optional RightJustify As Boolean) As String

Dim Pad, sl As Long
    
    sl = Len(Str)
    If sl > StrLen Then
        Pad = 0
        StringPad = Mid(Str, 1, StrLen)
    Else
        Pad = StrLen - sl
        If RightJustify = False Then
            StringPad = Str & String(Pad, " ")
        Else
            StringPad = String(Pad, " ") & Str
        End If
    End If

End Function

Private Sub TextParse(ByVal txt As String, ByVal SS As Byte)

Dim Row, Col As Byte
    
    For Row = 1 To 5
        SoldShip(SS, Row) = ""
    Next Row
    
    Row = 1
    x = ""
    For Col = 1 To Len(txt)
        y = Mid(txt, Col, 1)
        If y = vbCr Then
            SoldShip(SS, Row) = x
            Row = Row + 1
            If Row = 6 Then
                x = ""
                Exit For
            End If
            Col = Col + 1
            If Col = Len(txt) Then Exit For
            x = ""
        Else
            x = x & y
        End If
    Next Col
    If x <> "" Then
        SoldShip(SS, Row) = x
    End If

End Sub

Public Sub MasterItemUpd(ByVal QBID As String, _
                         ByVal JobID As Long, _
                         ByVal sQBName As String, _
                         ByVal sDescription As String, _
                         ByVal nCost As Currency, _
                         ByVal nPrice As Currency, _
                         ByVal nActive As Boolean, _
                         ByVal Inventory As Boolean)
                         
    ' add / change items for master list
    If JobID <> 0 Then Exit Sub
    
    SQLString = "SELECT * FROM InvStock WHERE JobID = 0 AND QBID = '" & QBID & "'"
    If InvStock.GetBySQL(SQLString) = False Then
        InvStock.Clear
        InvStock.QBID = QBID
        InvStock.JobID = 0
        InvStock.rsAdd
    End If
    
    InvStock.QBName = sQBName
    InvStock.Description = sDescription
    InvStock.Cost = nCost
    InvStock.MasterPrice = nPrice
    InvStock.CustomerPrice = nPrice
    InvStock.Active = nActive
    InvStock.InventoryItem = Inventory
    InvStock.rsPut

End Sub

Public Sub ItemUpd(ByVal JobID As Long)


    Dim InvStk0 As New cInvStock
    Dim InvStkJob As New cInvStock
    Dim AddFlag As Boolean
    Dim ChangeFlag As Boolean

    ' update the stock items for the job
    ' from the master list
    SQLString = "SELECT * FROM InvStock WHERE JobID = 0"
    
    ' wtf
    If InvStk0.GetBySQL(SQLString) = False Then Exit Sub
    
    Do

        AddFlag = False
        ChangeFlag = False
    
        SQLString = "SELECT * FROM InvStock WHERE JobID = " & JobID & _
                    " AND QBID = '" & InvStk0.QBID & "'"
        If InvStkJob.GetBySQL(SQLString) = False Then
            InvStkJob.Clear
            InvStkJob.QBID = InvStk0.QBID
            InvStkJob.QBName = InvStk0.QBName
            InvStkJob.JobID = JobID
            InvStkJob.StockSelect = True
            InvStkJob.CustomerPrice = InvStk0.MasterPrice
            AddFlag = True
        End If
    
        ' has anything changed?
        If InvStkJob.Description <> InvStk0.Description Then
            InvStkJob.Description = InvStk0.Description
            ChangeFlag = True
        End If
    
        If InvStkJob.Cost <> InvStk0.Cost Then
            InvStkJob.Cost = InvStk0.Cost
            ChangeFlag = True
        End If
    
        If InvStkJob.MasterPrice <> InvStk0.MasterPrice Then
            InvStkJob.MasterPrice = InvStk0.MasterPrice
            ChangeFlag = True
        End If
    
        If InvStkJob.Active <> InvStk0.Active Then
            InvStkJob.Active = InvStk0.Active
            ChangeFlag = True
        End If
    
        If InvStkJob.InventoryItem <> InvStk0.InventoryItem Then
            InvStkJob.InventoryItem = InvStk0.InventoryItem
            ChangeFlag = True
        End If
    
        If ChangeFlag = True Or AddFlag = True Then
            If AddFlag = True Then
                InvStkJob.rsAdd
            Else
                InvStkJob.rsPut
            End If
        End If

        If InvStk0.GetNext = False Then Exit Do
    
    Loop

End Sub

Public Function StripCtrlChar(ByVal InString As String) As String

    Dim stripI As Integer
    
    StripCtrlChar = ""
    
    For stripI = 1 To Len(InString)
        If Asc(Mid(InString, stripI, 1)) >= 32 Then
            StripCtrlChar = StripCtrlChar & Mid(InString, stripI, 1)
        End If
    Next stripI

End Function
