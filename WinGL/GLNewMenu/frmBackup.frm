VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackup 
   Caption         =   "Backup And Restore"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12960
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
   ScaleHeight     =   9435
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   6600
      TabIndex        =   5
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton cmdBU 
      Caption         =   "&Back Up"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   8520
      Width           =   1815
   End
   Begin VSFlex8Ctl.VSFlexGrid fgDBs 
      Height          =   4935
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   12375
      _cx             =   21828
      _cy             =   8705
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdSelectPath 
      Height          =   495
      Left            =   7800
      Picture         =   "frmBackup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtFolderName 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblMsg1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   11175
   End
   Begin VB.Label lblLastBU 
      Caption         =   "Last Backup ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   360
      Width           =   10455
   End
   Begin VB.Label Label2 
      Caption         =   "Backup Folder:"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fso As Object
Dim trs As New ADODB.Recordset
Dim x, y, z As String
Dim i, j, k As Integer
Dim mmsg As String

Private Sub Form_Load()
    
    ' last back up info
    SQLString = "SELECT * FROM PRGlobal WHERE Typecode = " & _
                PREquate.GlobalTypeLastBackUp & " AND UserID = " & UserID
    If PRGlobal.GetBySQL(SQLString) Then
        Me.txtFolderName.text = PRGlobal.Var1
        Me.lblLastBU.Caption = "Last Back Up: " & PRGlobal.Var2
    Else
        Me.lblLastBU.Caption = ""
    End If

    mmsg = "Data Backup"
    PopFG
    
End Sub

Private Sub PopFG()

    ' temp record set to show checks for the batch
    trs.CursorLocation = adUseClient

    trs.Fields.Append "BackUp", adBoolean
    trs.Fields.Append "CompanyName", adVarChar, 80, adFldIsNullable
    trs.Fields.Append "DBName", adVarChar, 80, adFldIsNullable
    trs.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM GLCompany ORDER BY Name"
    If Not GLCompany.GetBySQL(SQLString) Then
        MsgBox "No Company Files found to Back Up!", vbExclamation, mmsg
        Unload Me
    End If
    
    Do
        trs.AddNew
        trs.Fields("BackUp") = True
        trs.Fields("CompanyName") = Mid(GLCompany.Name, 1, 80)
        trs.Fields("DBName") = Mid(GLCompany.FileName, 1, 80)
        trs.Update
        If Not GLCompany.GetNext() Then Exit Do
    Loop
    
    SetGrid trs, Me.fgDBs

    With Me.fgDBs
        .ColWidth(1) = 5500
        .ColWidth(2) = 5500
        '.ColWidth(3) = 0
    
        .SelectionMode = flexSelectionByRow
        ' .Editable = flexEDNone
        .AutoSearch = flexSearchFromTop
    End With

End Sub

Private Sub cmdBU_Click()

Dim fnm As String
Dim result As Integer

    ' ** Disconnect **
    On Error Resume Next
    cn.Close
    cnDes.Close
    On Error GoTo 0
    
    Dim FileCount As Integer
    Dim ErrCount As Integer
    FileCount = 0
    ErrCount = 0
    trs.MoveFirst
    Do While Not trs.EOF
        If trs.Fields("BackUp") Then
            result = CopyDB(trs.Fields("DBName"))
            If result = 1 Then
                FileCount = FileCount + 1
            ElseIf result = -1 Then
                ErrCount = ErrCount + 1
            End If
        End If
        trs.MoveNext
    Loop
    
    ' >>> GLSystem
    result = CopyDB("\GLSystem.mdb")
    If result = 1 Then
        ' FileCount = FileCount + 1
    ElseIf result = -1 Then
        ErrCount = ErrCount + 1
    End If
    
    x = "Backup Complete - Company Files Backed Up: " & FileCount
    If ErrCount <> 0 Then x = x & vbCr & "Errors: " & ErrCount
    MsgBox x, vbInformation, mmsg

    ' ** Re-Connect **
    If Not CNOpen(dbCompany, dbPwd) Then
        MsgBox "Error opening company file: " & vbCr & dbCompany, vbExclamation, mmsg
        End
    End If
    
    If Not CNDesOpen(dbSystem) Then
        MsgBox "Error opening system file: " & vbCr & dbSystem, vbExclamation, mmsg
        End
    End If

    ' save last back up info
    SQLString = "SELECT * FROM PRGlobal WHERE Typecode = " & _
                PREquate.GlobalTypeLastBackUp & " AND UserID = " & UserID
    If Not PRGlobal.GetBySQL(SQLString) Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeLastBackUp
        PRGlobal.UserID = UserID
        PRGlobal.Save (Equate.RecAdd)
    End If
    PRGlobal.Var1 = Me.txtFolderName.text
    PRGlobal.Var2 = Date
    PRGlobal.Save (Equate.RecPut)
    
    Unload Me

End Sub

Function CopyDB(ByVal fnm As String) As Integer
    
    Dim FileExt As String
    Dim CopyFrom As String
    Dim CopyTo As String
    Dim OK As Boolean
    Dim GName As String
    Dim FromName As String
    Dim Pos As Integer
    
    
    Dim dbName As String
    dbName = mdbName(fnm)
    If NewADO Then
        dbName = Replace(LCase(dbName), ".mdb", ".accdb")
    End If
    
    If BalintFolder = "" Then
        CopyFrom = Left(App.Path, 1) & ":\Balint\Data\" & dbName
    Else
        CopyFrom = BalintFolder & "\Data\" & dbName
    End If
    CopyTo = AddBS(Me.txtFolderName) & dbName
    
    ' OK if DNE ...
    If Len(Dir(CopyFrom, vbNormal)) > 0 Then
        Me.lblMsg1.Caption = "Now Backing Up: " & dbName
        Me.Refresh
        On Error Resume Next
        FileCopy CopyFrom, CopyTo
        If Err.Number = 0 Then
            CopyDB = 1
        Else
            CopyDB = -1
            x = "Error copying " & CopyFrom & " to: " & CopyTo & vbCr & _
                Err.Number & " " & Err.Description
            MsgBox x, vbCritical, mmsg
        End If
        On Error GoTo 0
    Else
        CopyDB = 0
    End If

End Function

Private Sub go()
'    Dim Drv, fldr As String
'    Drv = Me.cmbDriveLetters.text
'    fldr = Drv & "\" & Me.txtFolderName
'    CreateFolder (fldr)
'    fldr = Drv & "\" & Me.txtFolderName & "\" & DTM_Stamp
'    CreateFolder (fldr)
'
'    MsgBox ("OK")

End Sub

Private Sub CreateFolder(ByVal fldr As String)
    If Not fso.FolderExists(fldr) Then
        MkDir fldr
    End If
End Sub

Private Sub DriveLetters()
    
'    Me.cmbDriveLetters.Clear
'    For I = 1 To 26
'        ' And LCase(Chr(96 + i)) <> "c"
'        If ValidDrive(Chr(96 + I)) = True Then
'           Me.cmbDriveLetters.AddItem UCase(LCase(Chr(96 + I))) + ": " + VolumeLabel(UCase(LCase(Chr(96 + I))))
'        End If
'    Next I
'    Me.cmbDriveLetters.ListIndex = 0
'    Me.txtFolderName = "Backup"

End Sub
    
Function VolumeLabel(drive As String) As String
    Dim Temp As String
    On Error Resume Next
    Temp = Dir$(drive, vbVolume)
    ' remove the period after the eigth character
    VolumeLabel = Left$(Temp, 8) + Mid$(Temp, 9)
End Function


Function ValidDrive(D As String) As Boolean
    On Error GoTo driveerror
    Dim Temp As String
    Temp = CurDir
    ChDrive D
    ChDir Temp
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

Private Sub cmdClearAll_Click()
    trs.MoveFirst
    Do While Not trs.EOF
        trs.Fields("BackUp") = False
        trs.Update
        trs.MoveNext
    Loop
    trs.MoveFirst
End Sub

Private Sub cmdSelectAll_Click()
    trs.MoveFirst
    Do While Not trs.EOF
        trs.Fields("BackUp") = True
        trs.Update
        trs.MoveNext
    Loop
    trs.MoveFirst
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSelectPath_Click()

   Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir    'Remember the current active directory
    CommonDialog1.DialogTitle = "Select a directory" 'titlebar
    CommonDialog1.InitDir = App.Path 'start dir, might be "C:\" or so also
    CommonDialog1.FileName = "Select a Directory"  'Something in filenamebox
    CommonDialog1.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog1.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    CommonDialog1.CancelError = True 'allow escape key/cancel
    CommonDialog1.ShowSave   'show the dialog screen

    If Err <> 32755 Then    ' User didn't chose Cancel.
        Me.txtFolderName = CurDir
    End If

    ' ChDir sTempDir  'restore path to what it was at entering

End Sub

