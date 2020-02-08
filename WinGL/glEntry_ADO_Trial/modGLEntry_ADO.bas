Attribute VB_Name = "modGLEntry_ADO"
Public Sub FFColumnCreate()
    ' for compatibility only
End Sub

Public Function ShowDate(ByVal thisDate As Date) As String
    ShowDate = Format(thisDate, "mm/dd/yyyy")
End Function

Public Function ShowValue(ByVal Amount As Currency) As String
    ShowValue = FormatCurrency(Amount, 2)
End Function

