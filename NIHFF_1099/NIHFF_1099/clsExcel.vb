Imports Excel = Microsoft.Office.Interop.Excel

Public Class clsExcel

    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Public xlname As String
    Dim SheetCount As Integer = 1
    Dim trow As Integer = 1

    Public Sub New()

        xl = New Excel.Application
        xl.Visible = True
        xl.DisplayAlerts = True
        wb = xl.Workbooks.Add()

        ' init summary tab
        ws = wb.Sheets(1)
        ws.Name = "Totals"
        ws.Cells(1, 1).value = "File Name"
        ws.Cells(1, 2).value = "Form Count"
        ws.Cells(1, 3).value = "Amount Total"

    End Sub

    Public Sub WriteTab(dt As DataTable, fnm As String)

        SheetCount += 1
        If SheetCount > wb.Sheets.Count Then
            wb.Sheets.Add(, xl.Sheets(xl.Sheets.Count))
        End If
        ws = wb.Sheets(SheetCount)
        ws.Name = fnm

        ' header row
        Dim colnum As Integer = 1
        Dim rownum As Integer = 1
        For Each dc As DataColumn In dt.Columns
            If dc.ColumnName <> "AmountLine" Then
                ws.Cells(1, colnum).value = dc.ColumnName
                colnum += 1
            End If
        Next

        Dim amttl As Double = 0
        For Each rw As DataRow In dt.Rows
            colnum = 1
            rownum += 1
            For Each dc As DataColumn In dt.Columns
                If dc.ColumnName <> "AmountLine" Then
                    If dc.ColumnName = "Amount" Then amttl += rw.Item("Amount")
                    ws.Cells(rownum, colnum).value = rw.Item(dc.ColumnName)
                    colnum += 1
                End If
            Next
        Next

        ws.Range("A1:Z1").EntireColumn.AutoFit()

        ' update totals tab
        trow += 1
        ws = wb.Sheets(1)
        ws.Cells(trow, 1).value = fnm
        ws.Cells(trow, 2).value = rownum - 1
        ws.Cells(trow, 3).value = amttl
        ws.Range("A1:Z1").EntireColumn.AutoFit()

    End Sub

End Class
