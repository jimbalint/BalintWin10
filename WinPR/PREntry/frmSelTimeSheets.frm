VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSelTimeSheets 
   Caption         =   "Select Time Sheets Entered"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
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
   ScaleHeight     =   7500
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT - DON'T USE TIME SHEETS"
      Height          =   615
      Left            =   4050
      TabIndex        =   2
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4335
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   6495
      _cx             =   11456
      _cy             =   7646
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
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   6375
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmSelTimeSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Public rsTimeSheet As New ADODB.Recordset
Public UseDist As Boolean

Dim i, j, k As Long
Dim X, Y, Z As String

Dim WE_Used As Boolean
Dim rsWE As New ADODB.Recordset
Dim GlobID, BID As Long
Dim StartWEDate As Long

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    
    ' escape disabled
    ' Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyEscape: cmdExit_Click
'    End Select
End Sub

Private Sub cmdExit_Click()
    OK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    ' sweep to fill in PRDist EmployerItemID missing
    ' possible problem with PRDE using timesheet data
    SQLString = "SELECT * FROM PRDist WHERE DistType = " & PREquate.ItemTypeOE & _
                " AND EmployerItemID = 0"
    If PRDist.GetBySQL(SQLString) = True Then
        Do
            If PRItem.GetByID(PRDist.ItemID) Then
                PRDist.EmployerItemID = PRItem.EmployerItemID
                PRDist.Save (Equate.RecPut)
            End If
            If PRDist.GetNext = False Then Exit Do
        Loop
    End If
    
    OK = True
    UseDist = False
    If rsTimeSheet.RecordCount > 0 Then
        rsTimeSheet.MoveFirst
        Do
            If rsTimeSheet!Selected = True Then
                UseDist = True
                Exit Do
            End If
            rsTimeSheet.MoveNext
        Loop Until rsTimeSheet.EOF
    End If
    
    If GlobID = 0 Then
        ' create new PRGlobal record
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypePRBatchWE
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Description = CStr(PRBatch.BatchID)
        PRGlobal.Save (Equate.RecAdd)
        GlobID = PRGlobal.GlobalID
    End If
    
    i = 0
    If PRGlobal.GetByID(GlobID) And rsTimeSheet.RecordCount > 0 Then
        rsTimeSheet.MoveFirst
        Do
            If rsTimeSheet!Selected = True Then
                i = i + 1
                If i = 11 Then
                    MsgBox "No more than 10 time sheet records allowed!", vbExclamation
                    Exit Sub
                End If
                If i = 1 Then PRGlobal.Var1 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 2 Then PRGlobal.Var2 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 3 Then PRGlobal.Var3 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 4 Then PRGlobal.Var4 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 5 Then PRGlobal.Var5 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 6 Then PRGlobal.Var6 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 7 Then PRGlobal.Var7 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 8 Then PRGlobal.Var8 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 9 Then PRGlobal.Var9 = CStr(CLng(rsTimeSheet!WEDate))
                If i = 10 Then PRGlobal.Var10 = CStr(CLng(rsTimeSheet!WEDate))
            End If
            rsTimeSheet.MoveNext
        Loop Until rsTimeSheet.EOF
    End If
    
    ' clear out the other fields
    For j = i + 1 To 10
        If j = 1 Then PRGlobal.Var1 = ""
        If j = 2 Then PRGlobal.Var2 = ""
        If j = 3 Then PRGlobal.Var3 = ""
        If j = 4 Then PRGlobal.Var4 = ""
        If j = 5 Then PRGlobal.Var5 = ""
        If j = 6 Then PRGlobal.Var6 = ""
        If j = 7 Then PRGlobal.Var7 = ""
        If j = 8 Then PRGlobal.Var8 = ""
        If j = 9 Then PRGlobal.Var9 = ""
        If j = 10 Then PRGlobal.Var10 = ""
    Next j
    
    PRGlobal.Save (Equate.RecPut)
    
    OK = True
    Me.Hide

End Sub

Public Sub Init()
    
    UseDist = False
    StartWEDate = 0
    
    ' define the record sets
    On Error Resume Next
    rsTimeSheet.Close
    On Error GoTo 0
    rsTimeSheet.CursorLocation = adUseClient
    rsTimeSheet.Fields.Append "Selected", adBoolean
    rsTimeSheet.Fields.Append "StartDate", adDate
    rsTimeSheet.Fields.Append "WEDate", adDate
    rsTimeSheet.Fields.Append "Hours", adSingle
    rsTimeSheet.Open , , adOpenDynamic, adLockOptimistic

    ' record set of WE dates already linked to a batch
    On Error Resume Next
    rsWE.Close
    On Error GoTo 0
    rsWE.CursorLocation = adUseClient
    rsWE.Fields.Append "WEDate", adDate
    rsWE.Fields.Append "CurrBatch", adBoolean
    rsWE.Open , , adOpenDynamic, adLockOptimistic

    ' get Time Sheet weeks ended already associated with a Batch
    BID = PRBatch.BatchID   ' store the current batch id
    GlobID = 0
    SQLString = "SELECT * FROM PRGlobal WHERE " & _
                "TypeCode = " & PREquate.GlobalTypePRBatchWE & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            
            ' BatchID store in PRGlobal.Description
            ' skip if not assigned
            If PRGlobal.Description = "" Then GoTo NextPRG
            
            ' does this batch still exists?
            ' user can delete batch to remove link to WE Date
            k = CLng(PRGlobal.Description)
            If PRBatch.GetByID(k) = False Then
                PRGlobal.Description = ""
                PRGlobal.Save (Equate.RecPut)
                GoTo NextPRG
            End If
            
            For i = 1 To 10
                If i = 1 Then X = PRGlobal.Var1
                If i = 2 Then X = PRGlobal.Var2
                If i = 3 Then X = PRGlobal.Var3
                If i = 4 Then X = PRGlobal.Var4
                If i = 5 Then X = PRGlobal.Var5
                If i = 6 Then X = PRGlobal.Var6
                If i = 7 Then X = PRGlobal.Var7
                If i = 8 Then X = PRGlobal.Var8
                If i = 9 Then X = PRGlobal.Var9
                If i = 10 Then X = PRGlobal.Var10

                If X <> "" Then
                    rsWE.AddNew
                    Z = CLng(PRGlobal.Description)
                    If Z = BID Then
                        ' used for current batch
                        rsWE!CurrBatch = True
                        GlobID = PRGlobal.GlobalID
                    Else
                        rsWE!CurrBatch = False
                    End If
                    rsWE!WEDate = CLng(X)
                    rsWE.Update
                                        
                    ' track the first WE date in use
                    If StartWEDate = 0 Or CLng(X) < StartWEDate Then
                        StartWEDate = CLng(X)
                    End If
                                        
                End If
            Next i
            
NextPRG:
            If PRGlobal.GetNext = False Then Exit Do
        Loop
    End If
            
    ' get the original PRBatch back
    If PRBatch.GetByID(BID) Then
    End If
    
    ' populate from PRTimeSheet - summ by WEDate
    
    ' patch for Hernandez - don't show WE date before it was stored in PRGlobal
    If StartWEDate = 0 Then
        SQLString = "SELECT * FROM PRTimeSheet"
    Else
        SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate > " & StartWEDate
    End If
    
    If PRTimeSheet.GetBySQL(SQLString) = False Then
        ' MsgBox "No Time Sheet records found", vbExclamation
        Me.Hide
        Exit Sub
    End If
    
    Do
        
        ' skip bogus data ?
        If PRTimeSheet.WEDate = 0 Then GoTo NextTS
        
        ' WE date already used - or in this batch?
        SQLString = "WEDate = " & PRTimeSheet.WEDate
        rsWE.Find SQLString, 0, adSearchForward, 1
        
        ' the WE date already used in another batch - don't show
        If rsWE.EOF = False Then
            If rsWE!CurrBatch = False Then GoTo NextTS
            WE_Used = True
        Else
            WE_Used = False
        End If
        
        If PRTimeSheet.TotalHours <> 0 Then
            SQLString = "WEDate  = " & PRTimeSheet.WEDate
            rsTimeSheet.Find SQLString, 0, adSearchForward, 1
            If rsTimeSheet.EOF Then
                rsTimeSheet.AddNew
                rsTimeSheet!WEDate = PRTimeSheet.WEDate
                rsTimeSheet!StartDate = PRTimeSheet.WEDate - 6
                rsTimeSheet!Hours = 0
                rsTimeSheet!Selected = WE_Used
            End If
            rsTimeSheet!Hours = rsTimeSheet!Hours + PRTimeSheet.TotalHours
            rsTimeSheet.Update
        End If
        
NextTS:
        If PRTimeSheet.GetNext = False Then Exit Do
    Loop

    rsTimeSheet.Sort = "WEDate DESC"

    SetGrid rsTimeSheet, fg

    With fg
        .TextMatrix(0, 1) = "Start Date"
        .TextMatrix(0, 2) = "EndDate"
        .TextMatrix(0, 3) = "Total Hours"
        .ColFormat(1) = "mm/dd/yyyy"
        .ColFormat(2) = "mm/dd/yyyy"
        .ColFormat(3) = "##,##0.00"
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1800
    End With

    If rsTimeSheet.RecordCount = 0 Then
        OK = False
        Me.UseDist = False
    Else
        Me.UseDist = True
    End If

End Sub
