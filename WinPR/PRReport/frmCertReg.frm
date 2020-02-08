VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCertReg 
   Caption         =   "Certified Payroll Register"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCertReg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRmvAllJob 
      Caption         =   "REMOVE ALL"
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelAllJob 
      Caption         =   "SELECT ALL"
      Height          =   495
      Left            =   9000
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdClrAllJob 
      Caption         =   "CLEAR ALL"
      Height          =   495
      Left            =   10560
      TabIndex        =   16
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdRmvAllEmp 
      Caption         =   "REMOVE ALL"
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelAllEmp 
      Caption         =   "SELECT ALL"
      Height          =   495
      Left            =   9120
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdClrAllEmp 
      Caption         =   "CLEAR ALL"
      Height          =   495
      Left            =   10680
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   4200
      TabIndex        =   12
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdRmvJob 
      Caption         =   "REMOVE"
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddJob 
      Caption         =   "ADD"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdRmvEmp 
      Caption         =   "REMOVE"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddEmp 
      Caption         =   "ADD"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEmp 
      Height          =   2415
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   9855
      _cx             =   17383
      _cy             =   4260
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
   Begin VB.ComboBox cmbWeekEnded 
      Height          =   360
      Left            =   6675
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   8640
      TabIndex        =   0
      Top             =   9000
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fgJob 
      Height          =   2415
      Left            =   2160
      TabIndex        =   8
      Top             =   6360
      Width           =   9855
      _cx             =   17383
      _cy             =   4260
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
   Begin VB.Label Label3 
      Caption         =   "Select Jobs:"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Select Employees:"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Select Week Ended:"
      Height          =   255
      Left            =   4275
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmCertReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, k As Long
Dim X, Y, Z As String
Dim boo As Boolean

Dim rsEmp As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsTS As New ADODB.Recordset
Dim rsHist As New ADODB.Recordset

Public LUType As String
Public SelID As Long
Dim JDate As Long
Dim cCert As caCert

' (1,x) = Job Totals
' (2,x) = Check totals
Dim Cert(2, 30) As Currency

Private Sub Form_Load()

    Set cCert = New caCert

    Me.lblCompanyName = PRCompany.Name

    CertSet rsEmp, fgEmp, PREquate.GlobalTypeUnionEmployee
    CertSet rsJob, fgJob, PREquate.GlobalTypeUnionJob
    
    ' populate the week ended combo
    JDate = Int(Now())
    
    ' find next Saturday
    Do
        If JDate Mod 7 = 0 Then Exit Do
        JDate = JDate + 1
    Loop
    
    With Me.cmbWeekEnded
        For i = 1 To 52
            .AddItem Format(JDate - 6, "mm/dd/yy") & " To: " & Format(JDate, "mm/dd/yy")
            .ItemData(.NewIndex) = JDate
            JDate = JDate - 7
        Next i
        .ListIndex = 0
    End With
    
    With frmCertLookUp
        
        .rs.CursorLocation = adUseClient
        .rs.Fields.Append "ID", adDouble
        .rs.Fields.Append "Select", adBoolean
        .rs.Fields.Append "Number", adDouble
        .rs.Fields.Append "Name", adVarChar, 40, adFldIsNullable
        .rs.Open , , adOpenDynamic, adLockOptimistic
        
        SetGrid .rs, .fg
        
        For i = 0 To .fg.Cols - 1
            .fg.ColKey(i) = .fg.TextMatrix(0, i)
        Next i
        
        .fg.ColHidden(.fg.ColIndex("ID")) = True
        .fg.ColWidth(.fg.ColIndex("Number")) = 2000
        .fg.ColWidth(.fg.ColIndex("Name")) = 7000
    
    End With
    
    ' rs to make sure use PRHist record only once
    With rsHist
        .CursorLocation = adUseClient
        .Fields.Append "HistID", adDouble
        .Open , , adOpenDynamic, adLockOptimistic
    End With
    
    Me.KeyPreview = True

End Sub

Private Sub cmdOK_Click()
    
    With Me.cmbWeekEnded
        JDate = .ItemData(.ListIndex)
    End With
    
    rsEmp.Filter = "Select = True"
    rsJob.Filter = "Select = True"
    
    rsJob.MoveFirst
    Do
        
        ' clear the array
        For i = 1 To 2
            For j = 1 To 31
                Cert(i, j) = 0
            Next j
        Next i
        
        ' clear the rs that keeps track of PRHist records used
        rsDelAll rsHist
        
        ' loop thru the employee records
        rsEmp.MoveFirst
        Do
            
            SQLString = "SELECT * FROM PRTimeSheet WHERE EmployeeID = " & rsEmp!ID & _
                        " AND JobID = " & rsJob!ID & _
                        " AND WEDate = " & JDate
            If PRTimeSheet.GetBySQL(SQLString) = True Then
                
                Do
                    
                    ' get the history record?
                    If PRTimeSheet.HistID = 0 Then
                        MsgBox "Time Sheet record for: " & rsEmp!Name & vbCr & _
                               " Job: " & rsJob!Name & " has not been processed!", vbExclamation
                        GoBack
                    End If
                    
                    rsHist.Find "HistID = " & PRTimeSheet.HistID, 0, adSearchForward, 1
                    If rsHist.EOF Then
                        
                        rsHist.AddNew
                        rsHist!HistID = PRTimeSheet.HistID
                        rsHist.Update
                    
                        ' get this history record
                        If PRHist.GetByID(PRTimeSheet.HistID) = False Then
                            MsgBox "PR History record for: " & rsEmp!Name & vbCr & _
                                   " Job: " & rsJob!Name & " not found!", vbExclamation
                            GoBack
                        End If
                        
                        ' add it to the array
                        CertUpdate 2, cCert.RegHrs, PRHist.RegHours
                        CertUpdate 2, cCert.RegHrs, PRHist.OEHours
                        CertUpdate 2, cCert.OvtHrs, PRHist.OTHours
                        CertUpdate 2, cCert.RegRate, PRHist.RegRate
                        CertUpdate 2, cCert.OvtRate, PRHist.OTRate
                        CertUpdate 2, cCert.RegGross, PRHist.RegAmount
                        CertUpdate 2, cCert.RegGross, PRHist.OEAmount
                        CertUpdate 2, cCert.OvtGross, PRHist.OTAmount
                        CertUpdate 2, cCert.Net, PRHist.Net
                        CertUpdate 2, cCert.SSTax, PRHist.SSTax
                        CertUpdate 2, cCert.MedTax, PRHist.MedTax
                        CertUpdate 2, cCert.FWTTax, PRHist.FWTTax
                        CertUpdate 2, cCert.SWTTax, PRHist.SWTTax
                        CertUpdate 2, cCert.CWTTax, PRHist.CWTTax
                        CertUpdate 2, cCert.TotalTax, PRHist.SSTax + PRHist.MedTax + PRHist.FWTTax _
                                      + PRHist.SWTTax + PRHist.CWTTax
                                      
                                      
                    End If
                    
                    ' update the timesheet amounts
                    If PRTimeSheet.ItemID = 99992 Then
                        CertUpdate 1, cCert.SunOvtHrs, PRTimeSheet.SunHours
                        CertUpdate 1, cCert.MonOvtHrs, PRTimeSheet.MonHours
                        CertUpdate 1, cCert.TueOvtHrs, PRTimeSheet.TueHours
                        CertUpdate 1, cCert.WedOvtHrs, PRTimeSheet.WedHours
                        CertUpdate 1, cCert.ThuOvtHrs, PRTimeSheet.ThuHours
                        CertUpdate 1, cCert.FriOvtHrs, PRTimeSheet.FriHours
                        CertUpdate 1, cCert.SatOvtHrs, PRTimeSheet.SatHours
                    Else
                        CertUpdate 1, cCert.SunRegHrs, PRTimeSheet.SunHours
                        CertUpdate 1, cCert.MonRegHrs, PRTimeSheet.MonHours
                        CertUpdate 1, cCert.TueRegHrs, PRTimeSheet.TueHours
                        CertUpdate 1, cCert.WedRegHrs, PRTimeSheet.WedHours
                        CertUpdate 1, cCert.ThuRegHrs, PRTimeSheet.ThuHours
                        CertUpdate 1, cCert.FriRegHrs, PRTimeSheet.FriHours
                        CertUpdate 1, cCert.SatRegHrs, PRTimeSheet.SatHours
                    End If
                    
                    If PRTimeSheet.GetNext = False Then Exit Do
                
                Loop
            
                ' calc aggregate amounts
                For i = 1 To 2
                    Cert(i, cCert.TotalTax) = Cert(i, cCert.SSTax) + Cert(i, cCert.MedTax) + Cert(i, cCert.FWTTax) _
                                              + Cert(i, cCert.SWTTax) + Cert(i, cCert.CWTTax)
                    Cert(i, cCert.TotGross) = Cert(i, cCert.RegGross) + Cert(i, cCert.OvtGross)
                    Cert(i, cCert.TotTaxGross) = Cert(i, cCert.RegTaxGross) + Cert(i, cCert.OvtTaxGross)
                Next i
        
        
        
        
        
                ' print the info for the job/employee
                
            End If
        
            rsEmp.MoveNext
            
        Loop Until rsEmp.EOF
        
        rsJob.MoveNext
    
    Loop Until rsJob.EOF

    rsEmp.Filter = adFilterNone
    rsJob.Filter = adFilterNone

End Sub

Private Sub CertUpdate(ByVal JobPay As Byte, ByVal Field As Byte, ByVal Amt As Currency)
    Cert(JobPay, Field) = Cert(JobPay, Field) + Amt
End Sub

Private Sub CertSet(ByRef rs As ADODB.Recordset, _
                    ByRef fg As VSFlexGrid, _
                    ByVal TypeCode As Byte)
                    
    With rs
        
        .CursorLocation = adUseClient
        .Fields.Append "ID", adDouble
        .Fields.Append "GlobalID", adDouble
        .Fields.Append "Select", adBoolean
        .Fields.Append "Number", adDouble
        .Fields.Append "Name", adVarChar, 40, adFldIsNullable
        .Open , , adOpenDynamic, adLockOptimistic
    
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & TypeCode & _
                    " AND UserID = " & PRCompany.CompanyID
        If PRGlobal.GetBySQL(SQLString) = True Then
            Do
                
                If TypeCode = PREquate.GlobalTypeUnionEmployee Then
                    If PREmployee.GetByID(PRGlobal.Var1) = True Then
                        .AddNew
                        !GlobalID = PRGlobal.GlobalID
                        !ID = PREmployee.EmployeeID
                        !Number = PREmployee.EmployeeNumber
                        !Name = Mid(PREmployee.LFName, 1, 40)
                        !Select = True
                        .Update
                    End If
                Else
                    If JCJob.GetByID(PRGlobal.Var1) = True Then
                        .AddNew
                        !GlobalID = PRGlobal.GlobalID
                        !ID = JCJob.JobID
                        !Number = JCJob.JobID
                        !Name = Mid(JCJob.FullName, 1, 40)
                        !Select = True
                        .Update
                    End If
                End If
                If PRGlobal.GetNext = False Then Exit Do
            Loop
        End If
    
        rs.Sort = "Name"
        
        SetGrid rs, fg
    
    End With
    
    With fg
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
        Next i
        .ColHidden(.ColIndex("ID")) = True
        .ColHidden(.ColIndex("GlobalID")) = True
        .ColWidth(.ColIndex("Number")) = 1500
        .ColWidth(.ColIndex("Name")) = 7400
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdAddEmp_Click()
    
    LUType = "emp"
    frmCertLookUp.Init
    frmCertLookUp.Show vbModal
    If frmCertLookUp.OK = False Then Exit Sub
    
    With frmCertLookUp.rs
 
        .Filter = "Select = True"
        
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do
        
            If PREmployee.GetByID(!ID) = True Then
                
                PRGlobal.Clear
                PRGlobal.UserID = PRCompany.CompanyID
                PRGlobal.TypeCode = PREquate.GlobalTypeUnionEmployee
                PRGlobal.Var1 = !ID
                PRGlobal.Save (Equate.RecAdd)
                
                rsEmp.AddNew
                rsEmp!GlobalID = PRGlobal.GlobalID
                rsEmp!ID = PREmployee.EmployeeID
                rsEmp!Name = Mid(PREmployee.LFName, 1, 40)
                rsEmp!Number = PREmployee.EmployeeNumber
                rsEmp!Select = True
                rsEmp.Update
    
            End If
            
            .MoveNext
            
        Loop Until .EOF
    
    End With
            
End Sub

Private Sub cmdRmvEmp_Click()

    If fgEmp.Row <= 0 Then Exit Sub
    
    SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & rsEmp!GlobalID
    cnDes.Execute SQLString
    
    rsEmp.Delete

End Sub

Private Sub cmdAddJob_Click()
    
    LUType = "job"
    frmCertLookUp.Init
    frmCertLookUp.Show vbModal
    If frmCertLookUp.OK = False Then Exit Sub
    
    With frmCertLookUp.rs
    
        .Filter = "Select = True"
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do
        
            If JCJob.GetByID(!ID) = True Then
            
                PRGlobal.Clear
                PRGlobal.UserID = PRCompany.CompanyID
                PRGlobal.TypeCode = PREquate.GlobalTypeUnionJob
                PRGlobal.Var1 = !ID
                PRGlobal.Save (Equate.RecAdd)
                
                rsJob.AddNew
                rsJob!GlobalID = PRGlobal.GlobalID
                rsJob!ID = JCJob.JobID
                rsJob!Name = Mid(JCJob.FullName, 1, 40)
                rsJob!Number = JCJob.JobID
                rsJob!Select = True
                rsJob.Update
    
            End If
            
            .MoveNext
            
        Loop Until .EOF
    
    End With
            
End Sub

Private Sub cmdRmvJob_Click()

    If fgJob.Row <= 0 Then Exit Sub
    
    SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & rsJob!GlobalID
    cnDes.Execute SQLString
    
    rsJob.Delete

End Sub

Private Sub cmdRmvAllEmp_Click()
    
    SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeUnionEmployee & _
                " AND UserID = " & PRCompany.CompanyID
    
    cnDes.Execute SQLString

    rsDelAll rsEmp

End Sub

Private Sub cmdRmvAllJob_Click()
    
    SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeUnionJob & _
                " AND UserID = " & PRCompany.CompanyID
    
    cnDes.Execute SQLString

    rsDelAll rsJob

End Sub


