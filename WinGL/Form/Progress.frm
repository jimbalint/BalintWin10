VERSION 5.00
Begin VB.Form frmProgress 
   Caption         =   "Windows PR - Program Progess"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   ControlBox      =   0   'False
   Icon            =   "Progress.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   6870
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   240
      Picture         =   "Progress.frx":030A
      ScaleHeight     =   675
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblMsg3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   9015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Balint Windows Accounting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   173
      TabIndex        =   1
      Top             =   2160
      Width           =   9015
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   173
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
