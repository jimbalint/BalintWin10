VERSION 5.00
Begin VB.Form frmBackup 
   Caption         =   "Backup And Restore"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFolderName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   3735
   End
   Begin VB.ComboBox cmbDriveLetters 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As Object

Private Sub Form_Load()
    Set fso = CreateObject("Scripting.FileSystemObject")
    MsgBox ("Under construction!!")
    Me.Hide
End Sub

Private Sub go()
    Dim Drv, fldr As String
    Drv = Me.cmbDriveLetters.text
    fldr = Drv & "\" & Me.txtFolderName
    CreateFolder (fldr)
    fldr = Drv & "\" & Me.txtFolderName & "\" & DTM_Stamp
    CreateFolder (fldr)
        
    MsgBox ("OK")

End Sub

Private Sub CreateFolder(ByVal fldr As String)
    If Not fso.FolderExists(fldr) Then
        MkDir fldr
    End If
End Sub

Private Sub DriveLetters()
    Me.cmbDriveLetters.Clear
    For I = 1 To 26
        ' And LCase(Chr(96 + i)) <> "c"
        If ValidDrive(Chr(96 + I)) = True Then
            Me.cmbDriveLetters.AddItem UCase(LCase(Chr(96 + I))) + ":"
        End If
    Next I
    Me.cmbDriveLetters.ListIndex = 0
    Me.txtFolderName = "Backup"
    
    

    'Dim f As New FileSystemObject, x As Drive, i As Long
'    Dim f As Object
'    Set f = CreateObject("Scripting.FileSystemObject")
'    MsgBox (f.drives.Count)
'    Exit Sub
    
'    ReDim Letters(f.drives.Count - 1)
'    For Each x In f.drives
'       Letters(i) = x.DriveLetter
'       MsgBox (x.DriveLetter)
'    Next x
'
'    Exit Sub
'
'
'    Dim drv As DriveListBox
'
'    Dim i As Long, Letters()
'    ReDim Letters(Drive1.ListCount - 1)
'    For i = 0 To Drive1.ListCount - 1
'       Letters(i) = Left$(Drive1.List(i), 1)
'       MsgBox (Letters(i))
'    Next i


End Sub

Function ValidDrive(D As String) As Boolean
    On Error GoTo driveerror
    Dim temp As String
    temp = CurDir
    ChDrive D
    ChDir temp
    ValidDrive = True
    Exit Function
driveerror:
End Function

Function DTM_Stamp() As String
    Dim dt As Date
    dt = Date
    DTM_Stamp = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2)
    DTM_Stamp = DTM_Stamp & "_" & Right("0" & Hour(Time), 2) & Right("0" & Minute(Time), 2) & Right("0" & Second(Time), 2)
End Function

