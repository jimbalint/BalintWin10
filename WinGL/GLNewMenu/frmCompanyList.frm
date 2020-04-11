VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCompanyList 
   Caption         =   "CLIENT LIST"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11055
      _cx             =   19500
      _cy             =   14208
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   11520
      TabIndex        =   2
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Default         =   -1  'True
      Height          =   615
      Left            =   11520
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Type the client name to locate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmCompanyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCompany As New ADODB.Recordset

Private Sub Form_Load()

    SQLString = "SELECT Name, FileName, ID FROM GLCompany ORDER BY Name"
    rsInit SQLString, cnDes, rsCompany
    
    If rsCompany.RecordCount = 0 Then
        MsgBox "No GL Company records found!", vbExclamation
        Unload Me
    End If
    
    SetGrid rsCompany, fg

    fg.ColWidth(0) = 5500
    fg.ColWidth(1) = 5500
    fg.ColWidth(2) = 0

    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    fg.AutoSearch = flexSearchFromTop

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    
Dim FName As String
    
    ' open the PRCompany database
    If BalintFolder = "" Then
        FName = Mid(App.Path, 1, 2) & Mid(rsCompany!FileName, 3, Len(rsCompany!FileName) - 2)
    Else
        FName = BalintFolder & "\Data\" & mdbName(rsCompany!FileName)
    End If
    If NewADO Then
        FName = Replace(FName, ".mdb", ".accdb")
    End If
    
    CNOpen FName, ""

    ' get the GLCompany record
    If Not GLCompany.GetData(rsCompany!ID) Then
        MsgBox "GLCompany Error! " & rsCompany!ID, vbExclamation
        End
    End If
    GLUser.LastCompany = GLCompany.ID

    ' get the PRCompany record
    If TableExists("PRCompany", cnDes) = True Then
        SQLString = "SELECT * FROM PRCompany WHERE GLCompanyID = " & rsCompany!ID
        If PRCompany.GetBySQL(SQLString) Then
            GLUser.LastPRCompany = PRCompany.CompanyID
        Else
            GLUser.LastPRCompany = 0
        End If
    End If
    
    ' update the last company fields of the user record
    GLUser.Save (Equate.RecPut)

    frmMainMenu.lblCompanyName = GLCompany.Name

    Unload Me

End Sub

Private Sub fg_DblClick()
    cmdSelect_Click
End Sub

Private Function TableExists(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim FString As String
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
           
    TableExists = False
           
    Do Until frs.EOF = True
                  
        If frs!Table_Name = TableName Then
            TableExists = True
            Exit Do
        End If
        
       frs.MoveNext
   
   Loop

End Function





