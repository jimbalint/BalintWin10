VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Print Selection"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
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
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Payroll Selection:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   4575
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton optWageReview 
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
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   3975
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
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
      End
      Begin VB.OptionButton optform941Print 
         Caption         =   "Form 941  -  Print"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   2655
      End
      Begin VB.OptionButton optForm941Entry 
         Caption         =   "Form 941  -  Entry"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label lbl941 
         Caption         =   "Form 941"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
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
      Left            =   3600
      TabIndex        =   8
      Top             =   3720
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
      frm941Entry.Hide
      frmWageReview.Hide
      frmStart.Hide
      frmLists.Show
   ElseIf optPRRpts = True Then
      frmLists.Hide
      frmRateList.Hide
      frm941Entry.Hide
      frmWageReview.Hide
      frmStart.Hide
      frmPRQtrlyRpts.Show
   ElseIf optRateList = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmStart.Hide
      frmWageReview.Hide
      frm941Entry.Hide
      frmRateList.Show
   ElseIf optWageReview = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmRateList.Hide
      frmStart.Hide
      frm941Entry.Hide
      frmWageReview.Show
   ElseIf optForm941Entry = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmStart.Hide
      frmRateList.Hide
      frmWageReview.Hide
      frm941Entry.Show
   ElseIf optform941Print = True Then
      frmLists.Hide
      frmPRQtrlyRpts.Hide
      frmStart.Hide
      frmRateList.Hide
      frmWageReview.Hide
      frm941Entry.Hide
      Form941Print
   End If

End Sub

