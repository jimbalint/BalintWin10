VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim x As String
Dim dte As Date

    dte = DateValue("04/13/2013")
    MsgBox DateSplit(dte)
    End
    
End Sub

Private Function SlashSplit(ByVal sString As String, sSide As Integer) As String
    ' divide string by slash - sSide = 1 - left of slash / sSide = 2 - right of slash
Dim sPos As Integer

    sPos = InStr(1, sString, "/")
    
    If sSide = 1 Then
        If sPos = 0 Then
            SlashSplit = sString
        Else
            SlashSplit = Mid(sString, 1, sPos - 1)
        End If
    Else
        If sPos = 0 Then
            SlashSplit = ""
        Else
            SlashSplit = Mid(sString, sPos + 1)
        End If
    End If

End Function

Private Function DateSplit(ByVal sDate As Date) As String
    DateSplit = Format(Month(sDate), "00") & "  " & Format(Day(sDate), "00") & "  " & Year(sDate)
End Function
