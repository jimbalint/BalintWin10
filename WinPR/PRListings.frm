VERSION 5.00
Begin VB.Form FileListing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Listings"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optName 
      Caption         =   "Name"
      Height          =   375
      Left            =   3255
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton optNumber 
      Caption         =   "Number"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame fraSort 
      Appearance      =   0  'Flat
      Caption         =   "Sort By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1695
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   555
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Payroll File Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "FileListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   GoBack
End Sub


