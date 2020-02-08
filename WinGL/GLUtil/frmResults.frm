VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Results"
   ClientHeight    =   9375
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   495
      Left            =   6150
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   10335
   End
   Begin VB.Label lblMsg3 
      Alignment       =   2  'Center
      Caption         =   "Msg3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   11175
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Msg2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   11175
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Msg1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   615
      TabIndex        =   4
      Top             =   840
      Width           =   11175
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   615
      TabIndex        =   3
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pg, i, j As Integer
Dim w, x, y, z As String
Dim MaxLn As Integer

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
   
   PrtInit ("Port")
   SetFont 10, Equate.Portrait
   Prvw.Caption = Me.Caption
    
   Ln = 0
   Pg = 0
   MaxLn = 55
    
   For j = 1 To List1.ListCount
       
       If Ln = 0 Or Ln > MaxLn Then Header
    
       Ln = Ln + 1
    
       PrintValue(1) = List1.List(j)
       FormatString(1) = "a80"
       FormatString(2) = "~"
       
       FormatPrint
   
   Next j

   Prvw.vsp.EndDoc
   Prvw.Show vbModal

End Sub


Private Sub Header()

   If Ln <> 0 Then FormFeed

   Ln = 0
   Pg = Pg + 1
   
   ' 29 characters for fixed left and right portion of first header line
   '    1             8       1   8                    10         1
   ' first line - system date & time / company name / page #
   x = Me.lblCompanyName
   y = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss")
   z = "Page: " & Format(Pg, "####")
   
   If Len(x) > Columns - 29 Then
      x = Mid(Me.Caption, 1, Columns - 29)
   End If
   
   ' center the company name in the string
   i = (Columns - 29 - Len(x)) / 2
   
   Ln = 1
   w = y & Space(i) & x & Space(i) & z
   PrtCenter Ln, w
   
   If Me.lblMsg1 <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, lblMsg1
   End If
   
   If Me.lblMsg2 <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, lblMsg2
   End If
   
   If Me.lblMsg3 <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, lblMsg3
   End If

   Ln = Ln + 1

End Sub



