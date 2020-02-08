Attribute VB_Name = "modGlobal"
' ------------------------------------------------------------------------------
' *** Direct Print Definitions ***
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

' ------------------------------------------------------------------------------
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
Public Sub DP_PrintLine(ByVal Str As String)
    Str = Str & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal Str, _
              Len(Str), lpcWritten)
End Sub



