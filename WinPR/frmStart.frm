VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Print Selection"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optRateList 
      Caption         =   "Payroll Rate File Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1448
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton optPRRpts 
      Caption         =   "Payroll Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton optPRListsLabels 
      Caption         =   "Payroll Lists and Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   6975
      Begin VB.OptionButton Option1 
         Caption         =   "Form 941 for 2008: Employer's QUARTERLY Federal Tax Return"
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
         Left            =   480
         TabIndex        =   8
         Top             =   2180
         Width           =   6375
      End
      Begin VB.OptionButton optSupplemental 
         Caption         =   "Payroll Report of Wages Supplemental"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1780
         Width           =   4215
      End
      Begin VB.OptionButton optWageReview 
         Caption         =   "Payroll Report of Wages Review"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5288
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   GoBack
End Sub

Private Sub cmdGo_Click()
   If optPRListsLabels = True Then
      frmPRQtrlyRpts.Hide
      frmRateList.Hide
      frm941.Hide
      frmWageReview.Hide
      frmStart.Hide
      frmLists.Show
   ElseIf optPRRpts = True Then
      frmLists.Hide
      frmRateList.Hide
      frm941.Hide
      frmWageReview.Hide
      frmStart.Hide
      frmPRQtrlyRpts.Show
   ElseIf optRateList = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmStart.Hide
      frmWageReview.Hide
      frm941.Hide
      frmRateList.Show
   ElseIf optWageReview = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmRateList.Hide
      frmStart.Hide
      frm941.Hide
      frmWageReview.Show
   ElseIf optForm941 = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmStart.Hide
      frmRateList.Hide
      frmWageReview.Hide
      frm941.Show
   End If

End Sub

Private Sub Form_Load()

End Sub
