Attribute VB_Name = "modQBData"
Option Explicit

Dim Flg As Boolean
Dim NewFlg As Boolean

Dim i, j, k As Long
Dim X, Y, Z As String

' *** QB Objects ***************************************

Dim companyQuery As ICompanyQuery
Dim requestMsgSet As IMsgSetRequest

Dim responseMsgSet As IMsgSetResponse
Dim ResponseList As IResponseList
Dim Response As IResponse

Dim companyRet As ICompanyRet

' Dim xmlMgr As New QBXMLRPLib.RequestProcessor2
Dim xmlMgr As New QBXMLRP2Lib.RequestProcessor2
' Dim xmlMgr As New QBXMLRPLib.RequestProcessor


Dim QBName, QBTicket As String

Public Function QBOpen(ByRef Frm As Form, ByVal Str As Label) As Boolean

    ' *** PRGlobal ***
    ' Var1 = 0 = Company Default / else UserID
    ' Var2 = QB FedID
    ' Var3 = QB CompanyName
    ' Var4 = QB FileName - Full Path
    ' Var5 = QB FileName - file name only
    
    ' if QB file already open
    '               compare to the company default PRGlobal record
    '               add entry for user if path not used before
    
    ' if QB file not open
    '               try company default first
    '               loop thru paths for user
    
    QBOpen = False
    
    ' ??????????????????????????????
    ' WTF ??? - just in case ???
    On Error Resume Next
    SessMgr.CloseConnection
    xmlMgr.CloseConnection
    On Error GoTo 0
    ' ??????????????????????????????
    
    ' does a default registration exist for this company?
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQB_Register & _
                " AND UserID = " & PRCompany.CompanyID & _
                " AND Var1 = '0'"
    If PRGlobal.GetBySQL(SQLString) = False Then
        MsgBox "This company is not registered with a QuickBooks File" & vbCr & vbCr & _
               "Please run the registration process", vbExclamation
        Exit Function
    End If
    
    Str = "Opening QuickBooks Connection a"
    Frm.Refresh
    
    ' is a QB file already open???
    ' xmlMgr.OpenConnection2 "", "Balint Accounting", localQBD
    xmlMgr.OpenConnection "", "Balint Accounting"
    
    Str = "Begin QuickBooks Session a"
    Frm.Refresh
    
    On Error Resume Next
    QBTicket = xmlMgr.BeginSession("", qbFileOpenDoNotCare)
        
    If Err.Number = 0 Then          ' QB file already opened
        
        On Error GoTo 0
        
        ' same file name?
        QBFileName = xmlMgr.GetCurrentCompanyFileName(QBTicket)
        QBName = GetFileName(QBFileName)        ' file name only - no path
        If Trim(PRGlobal.Var5) <> Trim(QBName) Then
            MsgBox "Incorrect QuickBooks file is open" & vbCr & vbCr & _
                   PRCompany.Name & " is registered to QuickBooks File: " & vbCr & vbCr & _
                   PRGlobal.Var4, vbExclamation
            xmlMgr.EndSession (QBTicket)
            xmlMgr.CloseConnection
            Exit Function
        End If
        
        xmlMgr.EndSession (QBTicket)
        xmlMgr.CloseConnection
        
        ' open in QBFC
        Str = "Opening QuickBooks Connection b"
        Frm.Refresh
        SessMgr.OpenConnection2 "", "Balint Accounting", ctLocalQBD
        
        Str = "Begin QuickBooks Session b"
        Frm.Refresh
        SessMgr.BeginSession "", omDontCare
        
        GetQBCompany
        
        ' check for Fed ID match
        If Trim(QBFedID) <> Trim(PRGlobal.Var2) Then
            Str = ""
            Frm.Refresh
            MsgBox "QuickBooks registered Federal ID is: " & PRGlobal.Var2 & vbCr & vbCr & _
                   "Currently openend QuickBooks file has: " & QBFedID & vbCr & vbCr & _
                   "Re-Run the QuickBooks registration process", vbExclamation
            SessMgr.EndSession
            SessMgr.CloseConnection
            Exit Function
        End If
        
        ' register the path for the user?
        Str = "Path registration review ..."
        Frm.Refresh
        If QBFileName <> PRGlobal.Var4 Then
            Flg = False
            SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQB_Register & _
                        " AND UserID = " & PRCompany.CompanyID & _
                        " AND Var1 = '" & User.ID & "'"
            If PRGlobal.GetBySQL(SQLString) Then
                Do
                    If PRGlobal.Var4 = QBFileName Then
                        Flg = True
                        Exit Do
                    End If
                    If PRGlobal.GetNext = False Then Exit Do
                Loop
            End If
            
            If Flg = False Then     ' add path for this user
                PRGlobal.Clear
                PRGlobal.TypeCode = PREquate.GlobalTypeQB_Register
                PRGlobal.UserID = PRCompany.CompanyID
                PRGlobal.Var1 = User.ID
                PRGlobal.Var2 = QBFedID
                PRGlobal.Var3 = QBCompanyName
                PRGlobal.Var4 = QBFileName
                PRGlobal.Var5 = QBName
                PRGlobal.Save (Equate.RecAdd)
            End If
        
        End If
        
        Str = ""
        Frm.Refresh
        
        QBOpen = True
        
        Exit Function
        
    ' no QB file open
    ' attempt from stored paths
    ElseIf Err.Number = PREquate.QBError_NoFileOpen _
       Or Err.Number = PREquate.QBError_QBBeginSession Then
            
        On Error GoTo 0
                   
        ' loop thru the files stored in PRGlobal - company default first
        Str = "Attempting stored QuickBooks Paths ..."
        Frm.Refresh
        
        QBName = ""
        Flg = False
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQB_Register & _
                    " AND UserID = " & PRCompany.CompanyID & _
                    " ORDER BY Var1"
        If PRGlobal.GetBySQL(SQLString) = False Then
            SessMgr.CloseConnection
            Str = ""
            Frm.Refresh
            Exit Function
        End If
        
        Do
            
            ' store just the file name from the company dflt
            If QBName = "" Then QBName = PRGlobal.Var5
            Str = "Attempting Open: " & PRGlobal.Var4
            Frm.Refresh
            
            On Error Resume Next
            GetAttr (PRGlobal.Var4)
            If Err.Number = 0 Then
                Flg = True
                Exit Do
            End If
            If PRGlobal.GetNext = False Then Exit Do
        
        Loop
        
        ' no files found
        If Flg = False Then
            MsgBox QBName & " has not been opened from this station yet" & vbCr & vbCr & _
                   "Please open the file in QuickBooks and try again", vbInformation
            SessMgr.CloseConnection
            Str = ""
            Frm.Refresh
            Exit Function
        End If
    
        ' attempt open to the file
        Str = "Open QB Connection ... 1"
        Frm.Refresh
        SessMgr.OpenConnection2 "", "Balint Accounting", ctLocalQBD
        
        On Error Resume Next
        
        Str = "Open QB Connection k1 ... 2"
        Frm.Refresh
        SessMgr.BeginSession PRGlobal.Var4, omDontCare
        
        If Err.Number = 0 Then
            On Error GoTo 0
            
            GetQBCompany
            
            If Trim(QBFedID) <> Trim(PRGlobal.Var2) Then
                MsgBox "QuickBooks registered Federal ID is: " & PRGlobal.Var2 & vbCr & vbCr & _
                       "Currently openend QuickBooks file has: " & QBFedID & vbCr & vbCr & _
                       "Re-Run the QuickBooks registration process", vbExclamation
                SessMgr.EndSession
                SessMgr.CloseConnection
                Exit Function
            
            End If
            
            Str = ""
            Frm.Refresh
            QBOpen = True
            Exit Function
        
        End If
        
        ' error ...
        
        MsgBox "QuickBooks Error: " & vbCr & vbCr & _
               Err.Number & vbCr & vbCr & _
               Err.Description, vbExclamation
               
        On Error GoTo 0
        
        SessMgr.CloseConnection
    
        Str = ""
        Frm.Refresh
    
    End If
                   
End Function

Public Function GetQBCompany() As Boolean

    GetQBCompany = False

    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    Set companyQuery = requestMsgSet.AppendCompanyQueryRq
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    If (responseMsgSet Is Nothing) Then
        MsgBox "Error gathering QB company data!", vbExclamation
        Exit Function
    End If

    Set ResponseList = responseMsgSet.ResponseList
    If (ResponseList Is Nothing) Then
        MsgBox "Error gathering QB company data!", vbExclamation
        Exit Function
    End If
    
    Set Response = ResponseList.GetAt(0)    ' only one company record
    Set companyRet = Response.Detail
    
    If (companyRet.EIN Is Nothing) Then
        QBFedID = ""
    Else
        QBFedID = companyRet.EIN.GetValue
    End If
    
    If (companyRet.CompanyName Is Nothing) Then
        QBCompanyName = ""
    Else
        QBCompanyName = companyRet.CompanyName.GetValue
    End If
    
    GetQBCompany = True

End Function


