VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String

Private Sub Form_Load()

    Ln = 1

'    Set Prvw = New frmPreview
'    Prvw.vsp.Preview = True
'
'    Prvw.vsp.MarginRight = 0
'    Prvw.vsp.MarginLeft = 0
'    Prvw.vsp.MarginBottom = 0
'    Prvw.vsp.MarginTop = 0
'
'
''    Prvw.vsp.Orientation = orPortrait
''    SetFont 10, Equate.Portrait
'
'    Prvw.vsp.StartDoc
'
''    Prt
'
''    Prvw.vsp.NewPage
'
'    Prvw.vsp.Orientation = orLandscape
'    SetFont 10, Equate.LandScape
'
'    Prt
    
    
    PrtInit "Port"
    ' SetFont 10, Equate.Portrait
    Prvw.vsp.Orientation = orLandscape
    Prt
    
    Prvw.vsp.Orientation = orLandscape
    Prvw.vsp.NewPage
    
    ' SetFont 10, Equate.LandScape
    Prt
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    
    End
   
End Sub

Private Sub Prt()

    For I = 1 To 10
        PrintValue(1) = "AAAAA":    FormatString(1) = "a10"
        PrintValue(2) = " ":        FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1
    Next I
    
End Sub
