Attribute VB_Name = "modNewField"
Option Explicit

Dim Lvl As Integer
Dim FWTRange(9), FWTAmount(9), FWTPct(9) As Currency
Dim SnglMarr As Byte
Dim MsgResponse As Variant
Dim boo As Boolean

Public Sub UpdateCheck(ByVal GLSys As Boolean, _
                       ByRef adoConn As ADODB.Connection)

Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim urs As ADODB.Recordset
Dim uCount As Long
Dim ucmd As ADODB.Command

Dim Ct1, Ct2, Recs As Long

    ' 2022-11-19
    If GLSys = True Then
        If AddField("PRFWTTable", "W4Type", "String", adoConn) Then
        End If
    End If

    ' 2012-01-08
    If GLSys = False Then
        If AddField("PRW2City", "Courtesy", "Byte", adoConn) Then
        End If
    End If
    
    ' for W2 processing
    If GLSys = True Then
        If AddField("GLCompany", "FederalID", "String", adoConn) Then
        End If
        If AddField("GLCompany", "SSN", "String", adoConn) Then
        End If
    End If

    ' ******** get these in first *******
    ' 02/13/2010 - add'l Dir Dep Fields
    If GLSys = True Then
        If AddField("PRCompany", "DirDepUseAltID", "Byte", adoConn) Then
        End If
        If AddField("PRCompany", "DirDepAltID", "Long", adoConn) Then
        End If
        If AddField("PRCompany", "GLCompanyID", "Long", adoConn) Then
        End If
                
        ' put a "1" before the FedID in NACHA file?
        If AddField("PRCompany", "DirDepID1", "Byte", adoConn) Then
        End If
    End If
    
    ' 02/08/2010
    If GLSys = True Then
        If AddField("PRCompany", "DirDepBalanced", "Byte", adoConn) Then
        End If
    End If

    ' upgrade from original Windows GL
    
    If GLSys = True Then
        If TableExists("PRGlobal", adoConn) = False Then
            GlobalCreate
        End If
        If TableExists("PRFWTTable", adoConn) = False Then
            FWTCreate
        End If
    End If

    ' 09/12/09 - Var fields in PRGlobal
    If GLSys = True Then
        If AddField("PRGlobal", "Month", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "UserID", "Long", adoConn) Then
        End If
        If AddField("PRGlobal", "Var1", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var2", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var3", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var4", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var5", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var6", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var7", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var8", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var9", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Var10", "String", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte1", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte2", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte3", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte4", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte5", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte6", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte7", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte8", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte9", "Byte", adoConn) Then
        End If
        If AddField("PRGlobal", "Byte10", "Byte", adoConn) Then
        End If
    End If

    ' GLSys = false - check for actual company data file
    ' GLSys = true - check for GLSystem File
        
    ' ******************************************************************
        
    If GLSys = False Then
        If TableExists("PRHist", adoConn) = False Then Exit Sub
        If TableExists("PRItem", adoConn) = False Then Exit Sub
    End If
        
    ' 09/27/2010
    If GLSys = False Then
        If TableExists("PRItem", cn) = True Then
            boo = AddField("PRItem", "CityID", "Long", adoConn)
        End If
    End If
    
    ' 08/09/2010
    If GLSys = False Then
        If TableExists("JCJob", cn) = True Then
            boo = AddField("JCJob", "Terms", "String", adoConn)
        End If
    End If
        
    ' 05/01/10 - PRHist update to QB Flag
    If GLSys = False Then
        boo = AddField("PRHist", "QBUpdateFlag", "Byte", adoConn)
    End If
    
    ' 04/12/10 - add qb invoice flag to PRDist
    If GLSys = False Then
        boo = AddField("PRDist", "QBInvoiceID", "String", adoConn)
    End If
        
    ' 04/28/10
    ' fix PRHist.StateID
    ' ** don't do for GLMenu - form progress non modal form err message **
    If GLSys = False And PRCompany.CompanyID <> 0 And InStr(1, UCase(App.EXEName), "GLMENU", vbTextCompare) = 0 Then
        If TableExists("PRHist", adoConn) = True Then
            SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & PRCompany.CompanyID & _
                        " AND Description = 'PRHIST STATEID FIX2'"
            If PRGlobal.GetBySQL(SQLString) = False Then
                Ct1 = 0
                Ct2 = 0
                frmProgress.Show
                frmProgress.lblMsg1 = PRCompany.Name
                frmProgress.lblMsg2 = "Now running PRHist.StateID Sweep ..."
                frmProgress.Refresh
                SQLString = "SELECT * FROM PRHist"
                If PRHist.GetBySQL(SQLString) = True Then
                    Recs = PRHist.Records
                    Do
                        Ct1 = Ct1 + 1
                        If Ct1 Mod 10 = 1 Then
                            frmProgress.lblMsg2 = "On Hist # " & Ct1 & " Of: " & Recs & " Fixed: " & Ct2
                            frmProgress.Refresh
                        End If
                        SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID
                        If PRDist.GetBySQL(SQLString) = True Then
                            If PRHist.StateID <> PRDist.StateID Then
                                PRHist.StateID = PRDist.StateID
                                PRHist.Save (Equate.RecPut)
                                Ct2 = Ct2 + 1
                            End If
                        End If
                        If PRHist.GetNext = False Then Exit Do
                    Loop
                End If
                PRGlobal.Clear
                PRGlobal.UserID = PRCompany.CompanyID
                PRGlobal.Description = "PRHIST STATEID FIX2"
                PRGlobal.Save (Equate.RecAdd)
                frmProgress.Hide
                If Ct2 > 0 Then
                    MsgBox Ct2 & " PR History records have been corrected" & vbCr & _
                           "Please verify any quarterly unemployment reports that have been run", _
                           vbInformation
                End If
            End If
        End If
    End If
        
    ' 04/19/10
    If GLSys = False Then
        boo = AddField("JCJob", "Active", "Byte", adoConn)
    End If
    
    ' 03/27/2010
    ' prevailing wage changes
    If GLSys = False Then
        boo = AddField("PRItem", "PWFringe", "Byte", adoConn)
        If TableExists("PRTimeSheet", adoConn) = True Then
            boo = AddField("PRTimeSheet", "PWCraftID", "Long", adoConn)
            boo = AddField("PRTimeSheet", "PWUnionID", "Long", adoConn)
            boo = AddField("PRTimeSheet", "PWRegRate", "Currency", adoConn)
            boo = AddField("PRTimeSheet", "PWOvtRate", "Currency", adoConn)
            boo = AddField("PRTimeSheet", "PWFringeAmt", "Currency", adoConn)
        End If
    End If
    
    ' 03/25/2010
    ' MD SWT Tables
    If GLSys = True Then
        SQLString = "SELECT * FROM PRFWTTable WHERE StateID = 21 AND TaxYear = 2010"
        If PRFWTTable.GetBySQL(SQLString) = False Then
            SWTMD2010Update
        End If
    End If
        
    ' 3/19/2010
    If GLSys = False Then
        boo = AddField("PRTimeSheet", "QBInvID", "String", adoConn)
    End If
    
    ' 3/13/2010
    If GLSys = False Then
        boo = AddField("JCCustomer", "QBTaxItem", "String", adoConn)
    End If
    
    If GLSys = False Then
        boo = AddField("PRTimeSheet", "BillingRate", "Currency", adoConn)
    End If
    
    ' 03/04/2010
    If GLSys = True Then
        boo = AddField("PRGlobal", "Byte1", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte2", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte3", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte4", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte5", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte6", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte7", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte8", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte9", "Byte", adoConn)
        boo = AddField("PRGlobal", "Byte10", "Byte", adoConn)
    End If
    
    ' 03/01/2010 - Rate Differential
    If GLSys = False Then
        If AddField("PRItem", "RateDifference", "Byte", adoConn) Then
        End If
    End If
    
    ' 02/27/2010 - Tax Code fields for QB Cust / Job
    If GLSys = False Then
        If TableExists("JCCustomer", adoConn) = True Then
            If AddField("JCCustomer", "QBTaxCode", "String", adoConn) Then
            End If
        End If
        If TableExists("JCJob", adoConn) = True Then
            If AddField("JCJob", "QBTaxCode", "String", adoConn) Then
            End If
        End If
    End If
    
    ' 02/06/10
    If GLSys = True Then
        If TableExists("PRCounty", adoConn) = False Then
            PRCountyCreate
        End If
    End If
    
    If GLSys = True Then
        If AddField("PRCity", "CountyID", "Long", adoConn) Then
        End If
    End If
    
    ' 02/01/10
    If GLSys = False Then
        If AddField("JCJob", "JobStatus", "Byte", adoConn) Then
        End If
    End If
    
    ' 01/26/10
    ' wage excluded from basis for deduct by percent (401k match purposes)
    If GLSys = False Then
        If AddField("PRItemHist", "WageExcluded", "Currency", adoConn) Then
        End If
    End If
    
    ' 01/18/10 - PRTimeSheet
    If GLSys = False Then
        If AddField("PRTimeSheet", "CustomerID", "Long", adoConn) Then
            ' fill it in
            SQLString = "SELECT * FROM PRTimeSheet"
            rsInit SQLString, cn, rs1
            If rs1.RecordCount > 0 Then
                Do
                    SQLString = "SELECT * FROM JCJob Where JobID = " & rs1!JobID
                    rsInit SQLString, cn, rs2
                    If rs2.RecordCount > 0 Then
                        rs1!CustomerID = rs2!ParentID
                        rs1.Update
                    End If
                    rs1.MoveNext
                Loop Until rs1.EOF
            End If
        End If
    End If
    
    ' 01/15/10 - New W2 Fields
    If GLSys = False Then
        If AddField("PRW2", "Void", "Byte", adoConn) Then
        End If
        If AddField("PRW2", "Skip", "Byte", adoConn) Then
        End If
    End If

    ' 01/02/10 - Sick Pay in PRItem
    If GLSys = False Then
        If AddField("PRItem", "SickPay", "Byte", adoConn) Then
        End If
    End If

    ' 01/29/2022 - OH eW2
    If GLSys = True Then
        boo = AddField("PRCompany", "OHeW2", "Byte", adoConn)
        If boo Then
            boo = AddField("PRCompany", "TermBiz", "Byte", adoConn)
            boo = AddField("PRCompany", "EstablishmentNumber", "String", adoConn)
            boo = AddField("PRCompany", "OtherEIN", "String", adoConn)
            boo = AddField("PRCompany", "KindOfEmployer", "String", adoConn)
            boo = AddField("PRCompany", "EmploymentCode", "String", adoConn)
            boo = AddField("PRCompany", "ThirdPartySickPay", "Byte", adoConn)
            boo = AddField("PRCompany", "ContactName", "String", adoConn)
            boo = AddField("PRCompany", "ContactPhoneNum", "String", adoConn)
            boo = AddField("PRCompany", "ContactPhoneExt", "String", adoConn)
            boo = AddField("PRCompany", "ContactFasNum", "String", adoConn)
            boo = AddField("PRCompany", "ContactEmail", "String", adoConn)
        End If
    End If

    ' 12/30/09 - 2010 tax update
    ' 12/31/10 - 2011 tax update
    ' 02/07/12 - 2012 tax update
    ' 01/13/13 - 2013 tax update
    If GLSys = False Then
        boo = AddField("PRHist", "MedAddAmt", "Currency", adoConn)
    End If
    
    If GLSys = True Then
        
        ' --- 2024 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2024"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2024
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 168600#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2024 updated to: $168,600", vbInformation
        End If
        
        ' --- 2023 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2023"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2023
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 160200#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2023 updated to: $160,200", vbInformation
        End If
        
        ' --- 2022 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2022"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2022
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 147000#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2022 updated to: $147,000", vbInformation
        End If
        
        ' --- 2023/11 and 2024 ***
        ' no more OH SWT multiplier!!!
        
        ' --- 2023 ***
        ' OH SWT multiplier
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeOHMultiplier & " " & _
                    "AND Year = 2023 and Month = 0"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeOHMultiplier
            PRGlobal.Year = 2023
            PRGlobal.Month = 0
            PRGlobal.Description = "OH Multiplier"
            PRGlobal.Amount = 1.001
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "OH Multiplier for Jan 2023 updated to: 1.001", vbInformation
        End If
        
        ' --- 2021 Sept ---------------------------------------------------------------
        ' OH SWT multiplier
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeOHMultiplier & " " & _
                    "AND Year = 2021 and Month = 9"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeOHMultiplier
            PRGlobal.Year = 2021
            PRGlobal.Month = 9
            PRGlobal.Description = "OH Multiplier"
            PRGlobal.Amount = 1.001
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "OH Multiplier for Sept 2021 updated to: 1.001", vbInformation
        End If
        
        ' --- 2021 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2021"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2021
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 142800#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2021 updated to: $142,800", vbInformation
        End If
        
        ' --- 2020 ---------------------------------------------------------------
        ' OH SWT multiplier
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeOHMultiplier & " " & _
                    "AND Year = 2019"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeOHMultiplier
            PRGlobal.Year = 2019
            PRGlobal.Description = "OH Multiplier"
            PRGlobal.Amount = 1.075
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "OH Multiplier for 2019 updated to: 1.075", vbInformation
        End If
        
        ' 2022 Revised W4
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTW4DepAmt & " " & _
                    "AND Year = 2022"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTW4DepAmt
            PRGlobal.Year = 2022
            PRGlobal.Description = "W4 Dependent Amt"
            PRGlobal.Amount = 2000
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "W4 Dependent Amt updated to $2,000.00", vbInformation
        End If
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTW4OtherDepAmt & " " & _
                    "AND Year = 2022"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTW4OtherDepAmt
            PRGlobal.Year = 2022
            PRGlobal.Description = "W4 Other Dependent Amt"
            PRGlobal.Amount = 500
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "W4 Other Dependent Amt updated to $500.00", vbInformation
        End If
                    
        ' 2021 - OH multiplier not changed
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeOHMultiplier & " " & _
                    "AND Year = 2020"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeOHMultiplier
            PRGlobal.Year = 2020
            PRGlobal.Description = "OH Multiplier"
            PRGlobal.Amount = 1.032
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "OH Multiplier for 2020 updated to: 1.032", vbInformation
        End If
                    
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
                    "AND Year = 2020"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
            PRGlobal.Year = 2020
            PRGlobal.Description = "FWT ALLOW"
            PRGlobal.Amount = 4300
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "FWT Allowance for 2020 updated to: $4,300", vbInformation
        End If
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2020"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2020
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 137700#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2020 updated to: $137,700", vbInformation
        End If
        
        ' --- 2019 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
                    "AND Year = 2019"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
            PRGlobal.Year = 2019
            PRGlobal.Description = "FWT ALLOW"
            PRGlobal.Amount = 4200
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "FWT Allowance for 2019 updated to: $4,200", vbInformation
        End If
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2019"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2019
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 132900#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2019 updated to: $132,900", vbInformation
        End If

        ' --- 2018 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
                    "AND Year = 2018"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
            PRGlobal.Year = 2018
            PRGlobal.Description = "FWT ALLOW"
            PRGlobal.Amount = 4150
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "FWT Allowance for 2018 updated to: $4,150", vbInformation
        End If
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2018"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2018
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 128400#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2018 updated to: $128,400", vbInformation
        End If

'        ' --- 2017 ---------------------------------------------------------------
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2017"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2017
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 4050
'            PRGlobal.Save (Equate.RecAdd)
'            MsgBox "FWT Allowance for 2017 updated to: $4,050", vbInformation
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2017"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2017
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 127200#
'            PRGlobal.Save (Equate.RecAdd)
'            MsgBox "SS Max for 2017 updated to: $127,200", vbInformation
'        End If
'
'        ' --- 2016 ---------------------------------------------------------------
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2016"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2016
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 4050
'            PRGlobal.Save (Equate.RecAdd)
'            MsgBox "FWT Allowance for 2016 updated to: $4,050", vbInformation
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2016"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2016
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 118500#
'            PRGlobal.Save (Equate.RecAdd)
'            MsgBox "SS Max for 2016 updated to: $118,500", vbInformation
'        End If

'        ' --- 2015 ---------------------------------------------------------------
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2015"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2015
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 4000
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2015"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2015
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 118500#
'            PRGlobal.Save (Equate.RecAdd)
'        End If

        ' --- 2015 ---------------------------------------------------------------


'        ' --- 2014 ---------------------------------------------------------------
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2014"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2014
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 3950
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2014"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2014
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 117000#
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        ' --- 2014 ---------------------------------------------------------------
'
'
'        ' --- 2013 ---------------------------------------------------------------
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2013"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2013
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 3900
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2013"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2013
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 113700#
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        ' --- 2013 ---------------------------------------------------------------


'        ' --- 2012 ---------------------------------------------------------------
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2012"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2012
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 3800
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2012"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2012
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 110100#
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        ' --- 2012 ---------------------------------------------------------------
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2011"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2011
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 3700
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2011"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2011
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 106800#
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
'                    "AND Year = 2010"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
'            PRGlobal.Year = 2010
'            PRGlobal.Description = "FWT ALLOW"
'            PRGlobal.Amount = 3650
'            PRGlobal.Save (Equate.RecAdd)
'        End If
'
'        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
'                    "AND Year = 2010"
'        If PRGlobal.GetBySQL(SQLString) = False Then
'            PRGlobal.Clear
'            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
'            PRGlobal.Year = 2010
'            PRGlobal.Description = "SS MAX"
'            PRGlobal.Amount = 106800#
'            PRGlobal.Save (Equate.RecAdd)
'        End If

        ' 2013-01-13
        ' med add pct & amt
        SQLString = " SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeMEDAddPct & " " & _
                    " AND Year = 2013"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeMEDAddPct
            PRGlobal.Year = 2013
            PRGlobal.Description = "MED Add Pct"
            PRGlobal.Amount = 0.9
            PRGlobal.Save (Equate.RecAdd)
        End If
        
        SQLString = " SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeMEDAddAmt & " " & _
                    " AND Year = 2013"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeMEDAddAmt
            PRGlobal.Year = 2013
            PRGlobal.Description = "MED Add Amt"
            PRGlobal.Amount = 200000
            PRGlobal.Save (Equate.RecAdd)
        End If
    
        ' Fed tax table update
' test for 2022 W4 tables
'MsgBox ("clear W4 2022...")
'SQLString = "delete * from PRFWTTable where W4Type <> '' and not isnull(W4Type) and TaxYear = 2022 and StateID = 0"
'cnDes.Execute SQLString
    
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2024 AND StateID = 0 and W4Type <> '' AND NOT ISNULL(W4Type)"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2024Update_W4
        
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2023 AND StateID = 0 and W4Type <> '' AND NOT ISNULL(W4Type)"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2023Update_W4
    
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2022 AND StateID = 0 and W4Type <> '' AND NOT ISNULL(W4Type)"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2022Update_W4
        
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2024 AND StateID = 0 and (W4Type = '' or ISNULL(W4Type))"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2024Update
        
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2023 AND StateID = 0 and (W4Type = '' or ISNULL(W4Type))"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2023Update
        
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2022 AND StateID = 0 and (W4Type = '' or ISNULL(W4Type))"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2022Update
        
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2021 AND StateID = 0 and (W4Type = '' or ISNULL(W4Type))"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2021Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2020 AND StateID = 0 and (W4Type = '' or ISNULL(W4Type))"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2020Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2019 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2019Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2018 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2018Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2017 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2017Update
        
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2016 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2016Update
        
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2015 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2015Update
        
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2014 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2014Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2013 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2013Update
        
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2012 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2012Update
        
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2011 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2011Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2010 AND StateID = 0"
'        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2010Update

        ' OH SWT table update
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2010 AND StateID = 36"
'        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2010Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2013 AND StateID = 36"
'        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2013Update
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2013 AND StateID = 36 AND TaxMonth = 9"
'        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2013UpdateSep1
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2014 AND StateID = 36 AND TaxMonth = 7"
'        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2014UpdateJul1
'
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2015 AND StateID = 36 AND TaxMonth = 8"
'        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2015UpdateAug1
    
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2019 AND StateID = 36 AND TaxMonth = 1"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2019Update
        
        ' 2023-11 eff - start in 2024
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2024 AND StateID = 36 AND TaxMonth = 1"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2024Update
    
    End If

    ' 12/12/2009 - comment fields / item memo
    If GLSys = False Then
        If TableExists("Notes", adoConn) = False Then
            NotesCreate
        End If
        If AddField("PRItem", "Comment", "String", adoConn) Then
        End If
        If AddField("PRItem", "DirDepRpt", "Byte", adoConn) Then
        End If
    End If

    ' 12/02/2009 - Employee Default Job ID
    If GLSys = False Then
        If AddField("PREmployee", "DefaultJobID", "Long", adoConn) Then
        End If
    End If

    ' 11/29/2009 - Job Cost fields
    If GLSys = False Then
        If AddField("PRBatch", "JobDist", "Byte", adoConn) Then
        End If
        If AddField("PRDist", "JobID", "Long", adoConn) Then
        End If
    End If

    ' 11/18/2009 - employee check comment
    If GLSys = False Then
        If AddField("PREmployee", "CheckComment", "String", adoConn) Then
        End If
    End If

    ' 11/03/2009 - dflt sort order
    If GLSys = True Then
        If AddField("PRCompany", "DfltSortOrder", "Byte", adoConn) Then
        End If
    End If

    ' 10/24/2009 - fields for NotInNet
    If GLSys = False Then
        If AddField("PRHist", "NotInNetAmount", "Currency", adoConn) Then
        End If
        If AddField("PRHist", "SDTax", "Currency", adoConn) Then
        End If
        If AddField("PRDist", "NotInNet", "Byte", adoConn) Then
        End If
        If AddField("PRDist", "SDTax", "Currency", adoConn) Then
        End If
    End If

    ' 10/23/2009 FiscalYear to GLPrint (for old installations)
    If GLSys = False Then
        If AddField("GLPrint", "FiscalYear", "Long", adoConn) Then
        End If
    End If

    ' 10/22/2009 Wkc Comp to PRHist
    If GLSys = False Then
        If AddField("PRHist", "WkcAmount", "Currency", adoConn) Then
        End If
    End If

    ' 10/20/09 - StateUnempID
    If GLSys = True Then
        If AddField("PRCompany", "StateUnempID", "String", adoConn) Then
        End If
    End If

    ' 09/15/09 - flag for Courtesy CWT add
    If GLSys = False Then
        If AddField("PREmployee", "CourtesyAdd", "Byte", adoConn) Then
        End If
    End If

    ' company phn #
    If GLSys = True Then
        If AddField("PRCompany", "PhoneNumber", "String", adoConn) Then
        End If
    End If

    ' 08/13/09 - Courtesy CWT
    If GLSys = False Then
        If AddField("PREmployee", "CourtesyCityID", "Long", adoConn) Then
        End If
        If AddField("PRDist", "CourtesyCityID", "Long", adoConn) Then
        End If
        If AddField("PRDist", "CourtesyCityTax", "Currency", adoConn) Then
        End If
        If AddField("PRDist", "ManualCourtesyCityTax", "Byte", adoConn) Then
        End If
    End If

    ' 08/03/09 - wage base to PRHist
    If GLSys = False Then
        If AddField("PRHist", "SSWageBase", "Currency", adoConn) Then
        End If
        If AddField("PRHist", "FUNWageBase", "Currency", adoConn) Then
        End If
        If AddField("PRHist", "SUNWageBase", "Currency", adoConn) Then
        End If
    End If

    ' 07/25/09 - add Unemployment max to PRState
    If GLSys = True Then
        If AddField("PRState", "UnempMax", "Currency", adoConn) Then
        End If
    End If

    ' 7/25/09 - add StateID to PRHist
    If GLSys = False Then
        If AddField("PRHist", "StateID", "Long", adoConn) Then
        End If
        If AddField("PRHist", "SUNWage", "Currency", adoConn) Then
        End If
    End If

    ' 01/22/08 - add LastPRCompany to GLCompany
    If GLSys = True Then
        If AddField("Users", "LastPRCompany", "Long", adoConn) Then
        End If
    End If
    
    ' 12/13/06 - add date/time posted field to GLHistory
    If GLSys = False Then
        If AddField("GLHistory", "PostDate", "DateTime", adoConn) <> 0 Then
           
           ' add a key also
           cn.Execute "CREATE INDEX PostKey ON GLHistory ([PostDate])"
           
           ' sweep in initial values
           rsInit "SELECT * FROM GLHistory ORDER BY ID", cn, urs
           
           ' display screen
           frmProgress.Show
           frmProgress.lblMsg1 = "Adding PostDate Field to GLHistory"
           
           Do Until urs.EOF
              
              urs.Fields("PostDate") = DateSerial(Year(Now()), Month(Now()), Day(Now())) + _
                                       TimeSerial(0, 0, urs!ID)
              urs.Update
              
              uCount = uCount + 1
              If uCount Mod 100 = 1 Then
                 frmProgress.lblMsg2 = "On Record: " & Format(uCount, "###,###,##0")
                 frmProgress.Refresh
              End If
              
              urs.MoveNext
           Loop
    
        End If

    End If

    Unload frmProgress

End Sub

Public Function AddField(ByVal TableName As String, _
                         ByVal ColumnName As String, _
                         ByVal ColumnType As String, _
                         ByRef adoConn As ADODB.Connection) _
                         As Byte
                         
Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim fString As String
Dim TblExists As Boolean
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
   
    FldFlag = False
    TblExists = False
    Do Until frs.EOF = True
          
        If UCase(frs!Table_Name) = UCase(TableName) Then
            TblExists = True
        End If
      
        If UCase(frs!Table_Name) = UCase(TableName) And UCase(frs!Column_Name) = UCase(ColumnName) Then
            FldFlag = True
            Exit Do
        End If
      
        frs.MoveNext
   
    Loop
    
    ' the table was not found
    If TblExists = False Then Exit Function
    
    ' the field already exists - no need to add it
    If FldFlag = True Then
        AddField = 0
        Exit Function
    End If
   
    frs.Close
    Set frs = Nothing
   
    ' add it - with retry
    Do
        
        On Error Resume Next
        
        fString = "ALTER TABLE " & TableName & _
                  " ADD COLUMN [" & ColumnName & "]" & _
                  " " & ColumnType
        adoConn.Execute fString
        
        If Err.Number = 0 Then
            AddField = 1
            Exit Do
        Else
            If InStr(1, LCase(Err.Description), "could not lock", vbTextCompare) Then
                MsgResponse = MsgBox("Database update not complete" & vbCr & _
                              "ALL other users must exit to proceed!", vbRetryCancel + vbExclamation)
                MsgBox TableName & "/" & ColumnName & " " & ColumnType & _
                     vbCrLf & vbCrLf & "Field Add Error" & Err.Description, _
                     vbOKOnly + vbCritical
                If MsgResponse = vbCancel Then
                    MsgBox "Update not complete - aborting ...", vbExclamation
                    End
                End If
            Else
                MsgBox TableName & "/" & ColumnName & " " & ColumnType & _
                     vbCrLf & vbCrLf & "Field Add Error" & Err.Description, _
                     vbOKOnly + vbCritical
                AddField = 0
                End
            End If
        End If
    
    Loop
    
End Function

Public Sub FieldSweep()

Dim DfltStateID As Long

    ' 02/10/2010 - move FUN pct
    If TableExists("PRGlobal", cn) = True Then
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeNudge & _
                    " AND Year >= 1900"
        If PRGlobal.GetBySQL(SQLString) = True Then
            Do
                PRGlobal.TypeCode = PREquate.GlobalTypeFUNPct
                PRGlobal.Save (Equate.RecPut)
                If PRGlobal.GetNext = False Then Exit Do
            Loop
        End If
    End If

    ' 08/03/09 - fill in PRHist.StateID if necessary
    If TableExists("PRHist", cn) = True Then
        SQLString = "SELECT * FROM PRHist WHERE IsNull(PRHist.StateID) OR " & _
                    " PRHist.StateID = 0"
        If PRHist.GetBySQL(SQLString) Then
            ' get the employer dflt city
            If Not PRCity.GetByID(PRCompany.DfltCityID) Then
                DfltStateID = 36    ' dflt to ohio
            Else
                If Not PRState.GetByID(PRCity.StateID) Then
                    DfltStateID = 36
                Else
                    DfltStateID = PRCity.StateID
                End If
            End If
            Do
                PRHist.StateID = DfltStateID
                PRHist.Save (Equate.RecPut)
                If Not PRHist.GetNext Then Exit Do
            Loop
        End If
    End If

End Sub
Private Sub FWT2014Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2250: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 11325: FWTAmount(3) = 907.5: FWTPct(3) = 15
            FWTRange(4) = 39150: FWTAmount(4) = 5081.25: FWTPct(4) = 25
            FWTRange(5) = 91600: FWTAmount(5) = 18193.75: FWTPct(5) = 28
            FWTRange(6) = 188600: FWTAmount(6) = 45353.75: FWTPct(6) = 33
            FWTRange(7) = 407350: FWTAmount(7) = 117541.25: FWTPct(7) = 35
            FWTRange(8) = 409000: FWTAmount(8) = 118118.75: FWTPct(8) = 39.6
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 8450: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 26600: FWTAmount(3) = 1815: FWTPct(3) = 15
            FWTRange(4) = 82250: FWTAmount(4) = 10162.5: FWTPct(4) = 25
            FWTRange(5) = 157300: FWTAmount(5) = 28925: FWTPct(5) = 28
            FWTRange(6) = 235300: FWTAmount(6) = 50765: FWTPct(6) = 33
            FWTRange(7) = 413550: FWTAmount(7) = 109587.5: FWTPct(7) = 35
            FWTRange(8) = 466050: FWTAmount(8) = 127962.5: FWTPct(8) = 39.6
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2014
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub
Private Sub FWT2024Update_W4()
    
    ' pub 15t MONTHLY tables
    Dim msh As Integer
    Dim twojob As Integer
    Dim tbltype As String
    Dim ftype
    ftype = Array("", "M", "S", "H")
    For msh = 1 To 3   ' 1 = Married / 2 = Single / 3 = "Head of Household"
        For twojob = 1 To 2
            If msh = 1 Then
                If twojob = 1 Then
                    ' FWT Married - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 2433: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 4367: FWTAmount(3) = 193.4: FWTPct(3) = 12
                    FWTRange(4) = 10292: FWTAmount(4) = 904.4: FWTPct(4) = 22
                    FWTRange(5) = 19188: FWTAmount(5) = 2861.52: FWTPct(5) = 24
                    FWTRange(6) = 34425: FWTAmount(6) = 6518.4: FWTPct(6) = 32
                    FWTRange(7) = 43054: FWTAmount(7) = 9279.68: FWTPct(7) = 35
                    FWTRange(8) = 63367: FWTAmount(8) = 16389.23: FWTPct(8) = 37
                Else
                    ' FWT Married - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1217: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 2183: FWTAmount(3) = 96.6: FWTPct(3) = 12
                    FWTRange(4) = 5146: FWTAmount(4) = 452.16: FWTPct(4) = 22
                    FWTRange(5) = 9594: FWTAmount(5) = 1430.72: FWTPct(5) = 24
                    FWTRange(6) = 17213: FWTAmount(6) = 3259.28: FWTPct(6) = 32
                    FWTRange(7) = 21527: FWTAmount(7) = 4639.76: FWTPct(7) = 35
                    FWTRange(8) = 31683: FWTAmount(8) = 8194.36: FWTPct(8) = 37
                End If
            ElseIf msh = 2 Then
                If twojob = 1 Then
                    ' FWT Single - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1217: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 2183: FWTAmount(3) = 96.6: FWTPct(3) = 12
                    FWTRange(4) = 5146: FWTAmount(4) = 452.16: FWTPct(4) = 22
                    FWTRange(5) = 9594: FWTAmount(5) = 1430.72: FWTPct(5) = 24
                    FWTRange(6) = 17213: FWTAmount(6) = 3259.28: FWTPct(6) = 32
                    FWTRange(7) = 21527: FWTAmount(7) = 4639.76: FWTPct(7) = 35
                    FWTRange(8) = 51996: FWTAmount(8) = 15303.91: FWTPct(8) = 37
                Else
                    ' FWT Single - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 608: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1092: FWTAmount(3) = 48.4: FWTPct(3) = 12
                    FWTRange(4) = 2573: FWTAmount(4) = 226.12: FWTPct(4) = 22
                    FWTRange(5) = 4797: FWTAmount(5) = 715.4: FWTPct(5) = 24
                    FWTRange(6) = 8606: FWTAmount(6) = 1629.56: FWTPct(6) = 32
                    FWTRange(7) = 10764: FWTAmount(7) = 2320.12: FWTPct(7) = 35
                    FWTRange(8) = 25998: FWTAmount(8) = 7652.02: FWTPct(8) = 37
                End If
            Else
                If twojob = 1 Then
                    ' FWT HOH - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1825: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 3204: FWTAmount(3) = 137.9: FWTPct(3) = 12
                    FWTRange(4) = 7083: FWTAmount(4) = 603.38: FWTPct(4) = 22
                    FWTRange(5) = 10200: FWTAmount(5) = 1289.12: FWTPct(5) = 24
                    FWTRange(6) = 17821: FWTAmount(6) = 3118.16: FWTPct(6) = 32
                    FWTRange(7) = 22133: FWTAmount(7) = 4498#: FWTPct(7) = 35
                    FWTRange(8) = 52604: FWTAmount(8) = 15162.85: FWTPct(8) = 37
                Else
                    ' FWT HOH - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 913: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1602: FWTAmount(3) = 68.9: FWTPct(3) = 12
                    FWTRange(4) = 3542: FWTAmount(4) = 301.7: FWTPct(4) = 22
                    FWTRange(5) = 5100: FWTAmount(5) = 644.46: FWTPct(5) = 24
                    FWTRange(6) = 8910: FWTAmount(6) = 1558.86: FWTPct(6) = 32
                    FWTRange(7) = 11067: FWTAmount(7) = 2249.1: FWTPct(7) = 35
                    FWTRange(8) = 26302: FWTAmount(8) = 7581.35: FWTPct(8) = 37
                End If
            End If
            
            tbltype = ftype(msh) & IIf(twojob = 2, "2", "")
        
            For Lvl = 1 To 8
    
                PRFWTTable.Clear
                PRFWTTable.TaxYear = 2024
                PRFWTTable.TaxMonth = 1
                PRFWTTable.StateID = 0
                PRFWTTable.W4Type = tbltype
    
                If Lvl = 1 Then
                    PRFWTTable.LowAmount = 0
                    PRFWTTable.ExcessBase = 0
                Else
                    PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                    PRFWTTable.ExcessBase = FWTRange(Lvl)
                End If
    
                If Lvl = 8 Then
                    PRFWTTable.HiAmount = 99999999.99
                Else
                    PRFWTTable.HiAmount = FWTRange(Lvl + 1)
                End If
    
                PRFWTTable.Amount = FWTAmount(Lvl)
                PRFWTTable.Percent = FWTPct(Lvl)
                PRFWTTable.Save (Equate.RecAdd)
    
            Next Lvl
        
        Next twojob
    Next msh

    MsgBox "Federal tax tables ** Revised W4 ** updated for 2024!", vbOKOnly + vbInformation

End Sub

Private Sub FWT2023Update_W4()
    
    ' pub 15t MONTHLY tables
    Dim msh As Integer
    Dim twojob As Integer
    Dim tbltype As String
    Dim ftype
    ftype = Array("", "M", "S", "H")
    For msh = 1 To 3   ' 1 = Married / 2 = Single / 3 = "Head of Household"
        For twojob = 1 To 2
            If msh = 1 Then
                If twojob = 1 Then
                    ' FWT Married - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 2308: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 4142: FWTAmount(3) = 183.4: FWTPct(3) = 12
                    FWTRange(4) = 9763: FWTAmount(4) = 857.92: FWTPct(4) = 22
                    FWTRange(5) = 18204: FWTAmount(5) = 2714.94: FWTPct(5) = 24
                    FWTRange(6) = 32658: FWTAmount(6) = 6183.9: FWTPct(6) = 32
                    FWTRange(7) = 40850: FWTAmount(7) = 8805.34: FWTPct(7) = 35
                    FWTRange(8) = 60121: FWTAmount(8) = 15550.19: FWTPct(8) = 37
                Else
                    ' FWT Married - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1154: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 2071: FWTAmount(3) = 91.7: FWTPct(3) = 12
                    FWTRange(4) = 4881: FWTAmount(4) = 428.9: FWTPct(4) = 22
                    FWTRange(5) = 9102: FWTAmount(5) = 1357.52: FWTPct(5) = 24
                    FWTRange(6) = 16329: FWTAmount(6) = 3092: FWTPct(6) = 32
                    FWTRange(7) = 20425: FWTAmount(7) = 4402.72: FWTPct(7) = 35
                    FWTRange(8) = 30060: FWTAmount(8) = 7774.97: FWTPct(8) = 37
                End If
            ElseIf msh = 2 Then
                If twojob = 1 Then
                    ' FWT Single - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1154: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 2071: FWTAmount(3) = 91.7: FWTPct(3) = 12
                    FWTRange(4) = 4881: FWTAmount(4) = 428.9: FWTPct(4) = 22
                    FWTRange(5) = 9102: FWTAmount(5) = 1357.52: FWTPct(5) = 24
                    FWTRange(6) = 16329: FWTAmount(6) = 3092: FWTPct(6) = 32
                    FWTRange(7) = 20425: FWTAmount(7) = 4402.72: FWTPct(7) = 35
                    FWTRange(8) = 49331: FWTAmount(8) = 14519.82: FWTPct(8) = 37
                Else
                    ' FWT Single - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 577: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1035: FWTAmount(3) = 45.8: FWTPct(3) = 12
                    FWTRange(4) = 2441: FWTAmount(4) = 214.52: FWTPct(4) = 22
                    FWTRange(5) = 4551: FWTAmount(5) = 678.72: FWTPct(5) = 24
                    FWTRange(6) = 8165: FWTAmount(6) = 1546.08: FWTPct(6) = 32
                    FWTRange(7) = 10213: FWTAmount(7) = 2201.44: FWTPct(7) = 35
                    FWTRange(8) = 24666: FWTAmount(8) = 7259.99: FWTPct(8) = 37
                End If
            Else
                If twojob = 1 Then
                    ' FWT HOH - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1733: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 3042: FWTAmount(3) = 130.9: FWTPct(3) = 12
                    FWTRange(4) = 6721: FWTAmount(4) = 572.38: FWTPct(4) = 22
                    FWTRange(5) = 9679: FWTAmount(5) = 1223.14: FWTPct(5) = 24
                    FWTRange(6) = 16908: FWTAmount(6) = 2958.1: FWTPct(6) = 32
                    FWTRange(7) = 21004: FWTAmount(7) = 4268.82: FWTPct(7) = 35
                    FWTRange(8) = 49908: FWTAmount(8) = 14385.22: FWTPct(8) = 37
                Else
                    ' FWT HOH - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 867: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1521: FWTAmount(3) = 65.4: FWTPct(3) = 12
                    FWTRange(4) = 3360: FWTAmount(4) = 286.08: FWTPct(4) = 22
                    FWTRange(5) = 4840: FWTAmount(5) = 611.68: FWTPct(5) = 24
                    FWTRange(6) = 8454: FWTAmount(6) = 1479.04: FWTPct(6) = 32
                    FWTRange(7) = 10502: FWTAmount(7) = 2134.4: FWTPct(7) = 35
                    FWTRange(8) = 24954: FWTAmount(8) = 7192.6: FWTPct(8) = 37
                End If
            End If
            
            tbltype = ftype(msh) & IIf(twojob = 2, "2", "")
        
            For Lvl = 1 To 8
    
                PRFWTTable.Clear
                PRFWTTable.TaxYear = 2023
                PRFWTTable.TaxMonth = 1
                PRFWTTable.StateID = 0
                PRFWTTable.W4Type = tbltype
    
                If Lvl = 1 Then
                    PRFWTTable.LowAmount = 0
                    PRFWTTable.ExcessBase = 0
                Else
                    PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                    PRFWTTable.ExcessBase = FWTRange(Lvl)
                End If
    
                If Lvl = 8 Then
                    PRFWTTable.HiAmount = 99999999.99
                Else
                    PRFWTTable.HiAmount = FWTRange(Lvl + 1)
                End If
    
                PRFWTTable.Amount = FWTAmount(Lvl)
                PRFWTTable.Percent = FWTPct(Lvl)
                PRFWTTable.Save (Equate.RecAdd)
    
            Next Lvl
        
        Next twojob
    Next msh

    MsgBox "Federal tax tables ** Revised W4 ** updated for 2023!", vbOKOnly + vbInformation

End Sub


Private Sub FWT2022Update_W4()
    
    ' pub 15t MONTHLY tables
    Dim msh As Integer
    Dim twojob As Integer
    Dim tbltype As String
    Dim ftype
    ftype = Array("", "M", "S", "H")
    For msh = 1 To 3   ' 1 = Married / 2 = Single / 3 = "Head of Household"
        For twojob = 1 To 2
            If msh = 1 Then
                If twojob = 1 Then
                    ' FWT Married - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 2158: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 3871: FWTAmount(3) = 171.3: FWTPct(3) = 12
                    FWTRange(4) = 9121: FWTAmount(4) = 801.3: FWTPct(4) = 22
                    FWTRange(5) = 17004: FWTAmount(5) = 2535.56: FWTPct(5) = 24
                    FWTRange(6) = 30500: FWTAmount(6) = 5774.6: FWTPct(6) = 32
                    FWTRange(7) = 38150: FWTAmount(7) = 8222.6: FWTPct(7) = 35
                    FWTRange(8) = 56146: FWTAmount(8) = 14521.2: FWTPct(8) = 37
                Else
                    ' FWT Married - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1079: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1935: FWTAmount(3) = 85.6: FWTPct(3) = 12
                    FWTRange(4) = 4560: FWTAmount(4) = 400.6: FWTPct(4) = 22
                    FWTRange(5) = 8502: FWTAmount(5) = 1267.84: FWTPct(5) = 24
                    FWTRange(6) = 15250: FWTAmount(6) = 2887.36: FWTPct(6) = 32
                    FWTRange(7) = 19075: FWTAmount(7) = 4111.36: FWTPct(7) = 35
                    FWTRange(8) = 28073: FWTAmount(8) = 7260.66: FWTPct(8) = 37
                End If
            ElseIf msh = 2 Then
                If twojob = 1 Then
                    ' FWT Single - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1079: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1935: FWTAmount(3) = 85.6: FWTPct(3) = 12
                    FWTRange(4) = 4560: FWTAmount(4) = 400.6: FWTPct(4) = 22
                    FWTRange(5) = 8502: FWTAmount(5) = 1267.84: FWTPct(5) = 24
                    FWTRange(6) = 15250: FWTAmount(6) = 2887.36: FWTPct(6) = 32
                    FWTRange(7) = 19075: FWTAmount(7) = 4111.36: FWTPct(7) = 35
                    FWTRange(8) = 46071: FWTAmount(8) = 13559.96: FWTPct(8) = 37
                Else
                    ' FWT Single - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 540: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 968: FWTAmount(3) = 42.8: FWTPct(3) = 12
                    FWTRange(4) = 2280: FWTAmount(4) = 200.24: FWTPct(4) = 22
                    FWTRange(5) = 4251: FWTAmount(5) = 633.86: FWTPct(5) = 24
                    FWTRange(6) = 7625: FWTAmount(6) = 1443.62: FWTPct(6) = 32
                    FWTRange(7) = 9538: FWTAmount(7) = 2055.78: FWTPct(7) = 35
                    FWTRange(8) = 23035: FWTAmount(8) = 6779.73: FWTPct(8) = 37
                End If
            Else
                If twojob = 1 Then
                    ' FWT HOH - 1 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 1617: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 2838: FWTAmount(3) = 122.1: FWTPct(3) = 12
                    FWTRange(4) = 6275: FWTAmount(4) = 534.54: FWTPct(4) = 22
                    FWTRange(5) = 9038: FWTAmount(5) = 1142.4: FWTPct(5) = 24
                    FWTRange(6) = 15788: FWTAmount(6) = 2762.4: FWTPct(6) = 32
                    FWTRange(7) = 19613: FWTAmount(7) = 3986.4: FWTPct(7) = 35
                    FWTRange(8) = 46608: FWTAmount(8) = 13434.65: FWTPct(8) = 37
                Else
                    ' FWT HOH - 2 job
                    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
                    FWTRange(2) = 808: FWTAmount(2) = 0: FWTPct(2) = 10
                    FWTRange(3) = 1419: FWTAmount(3) = 61.1: FWTPct(3) = 12
                    FWTRange(4) = 3138: FWTAmount(4) = 267.38: FWTPct(4) = 22
                    FWTRange(5) = 4519: FWTAmount(5) = 571.2: FWTPct(5) = 24
                    FWTRange(6) = 7894: FWTAmount(6) = 1381.2: FWTPct(6) = 32
                    FWTRange(7) = 9806: FWTAmount(7) = 1993.04: FWTPct(7) = 35
                    FWTRange(8) = 23304: FWTAmount(8) = 6717.34: FWTPct(8) = 37
                End If
            End If
            
            tbltype = ftype(msh) & IIf(twojob = 2, "2", "")
        
            For Lvl = 1 To 8
    
                PRFWTTable.Clear
                PRFWTTable.TaxYear = 2022
                PRFWTTable.TaxMonth = 1
                PRFWTTable.StateID = 0
                PRFWTTable.W4Type = tbltype
    
                If Lvl = 1 Then
                    PRFWTTable.LowAmount = 0
                    PRFWTTable.ExcessBase = 0
                Else
                    PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                    PRFWTTable.ExcessBase = FWTRange(Lvl)
                End If
    
                If Lvl = 8 Then
                    PRFWTTable.HiAmount = 99999999.99
                Else
                    PRFWTTable.HiAmount = FWTRange(Lvl + 1)
                End If
    
                PRFWTTable.Amount = FWTAmount(Lvl)
                PRFWTTable.Percent = FWTPct(Lvl)
                PRFWTTable.Save (Equate.RecAdd)
    
            Next Lvl
        
        Next twojob
    Next msh

    MsgBox "Federal tax tables ** Revised W4 ** updated for 2022!", vbOKOnly + vbInformation

End Sub
Private Sub FWT2023Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 5250: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 16250: FWTAmount(3) = 1100: FWTPct(3) = 12
            FWTRange(4) = 49975: FWTAmount(4) = 5147: FWTPct(4) = 22
            FWTRange(5) = 100625: FWTAmount(5) = 16290: FWTPct(5) = 24
            FWTRange(6) = 187350: FWTAmount(6) = 37104: FWTPct(6) = 32
            FWTRange(7) = 236500: FWTAmount(7) = 52832: FWTPct(7) = 35
            FWTRange(8) = 583375: FWTAmount(8) = 174238.25: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 14800: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 36800: FWTAmount(3) = 2200: FWTPct(3) = 12
            FWTRange(4) = 104250: FWTAmount(4) = 10294: FWTPct(4) = 22
            FWTRange(5) = 205550: FWTAmount(5) = 32580: FWTPct(5) = 24
            FWTRange(6) = 379000: FWTAmount(6) = 74208: FWTPct(6) = 32
            FWTRange(7) = 477300: FWTAmount(7) = 105664: FWTPct(7) = 35
            FWTRange(8) = 708550: FWTAmount(8) = 186601.5: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2023
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2023!", vbOKOnly + vbInformation

End Sub
Private Sub FWT2024Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 6000: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 17600: FWTAmount(3) = 1160: FWTPct(3) = 12
            FWTRange(4) = 53150: FWTAmount(4) = 5426: FWTPct(4) = 22
            FWTRange(5) = 106525: FWTAmount(5) = 17168.5: FWTPct(5) = 24
            FWTRange(6) = 197950: FWTAmount(6) = 39110.5: FWTPct(6) = 32
            FWTRange(7) = 249725: FWTAmount(7) = 55678.5: FWTPct(7) = 35
            FWTRange(8) = 615350: FWTAmount(8) = 183647.25: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 16300: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 39500: FWTAmount(3) = 2320: FWTPct(3) = 12
            FWTRange(4) = 110600: FWTAmount(4) = 10852: FWTPct(4) = 22
            FWTRange(5) = 217350: FWTAmount(5) = 34337: FWTPct(5) = 24
            FWTRange(6) = 400200: FWTAmount(6) = 78221: FWTPct(6) = 32
            FWTRange(7) = 503750: FWTAmount(7) = 111357: FWTPct(7) = 35
            FWTRange(8) = 747500: FWTAmount(8) = 196669.5: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2024
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2024!", vbOKOnly + vbInformation

End Sub

Private Sub FWT2022Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 4350: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 14625: FWTAmount(3) = 1027.5: FWTPct(3) = 12
            FWTRange(4) = 46125: FWTAmount(4) = 4807.5: FWTPct(4) = 22
            FWTRange(5) = 93425: FWTAmount(5) = 15213.5: FWTPct(5) = 24
            FWTRange(6) = 174400: FWTAmount(6) = 34647.5: FWTPct(6) = 32
            FWTRange(7) = 220300: FWTAmount(7) = 49335.5: FWTPct(7) = 35
            FWTRange(8) = 544250: FWTAmount(8) = 162718: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 13000: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 33550: FWTAmount(3) = 2055: FWTPct(3) = 12
            FWTRange(4) = 96550: FWTAmount(4) = 9615: FWTPct(4) = 22
            FWTRange(5) = 191150: FWTAmount(5) = 30427: FWTPct(5) = 24
            FWTRange(6) = 353100: FWTAmount(6) = 69295: FWTPct(6) = 32
            FWTRange(7) = 444900: FWTAmount(7) = 98671: FWTPct(7) = 35
            FWTRange(8) = 660850: FWTAmount(8) = 174253.5: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2022
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2022!", vbOKOnly + vbInformation

End Sub


Private Sub FWT2021Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 3950: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 13900: FWTAmount(3) = 995: FWTPct(3) = 12
            FWTRange(4) = 44475: FWTAmount(4) = 4664: FWTPct(4) = 22
            FWTRange(5) = 90325: FWTAmount(5) = 14751: FWTPct(5) = 24
            FWTRange(6) = 168875: FWTAmount(6) = 33603: FWTPct(6) = 32
            FWTRange(7) = 213375: FWTAmount(7) = 47843: FWTPct(7) = 35
            FWTRange(8) = 527550: FWTAmount(8) = 157804.25: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 12200: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 32100: FWTAmount(3) = 1990: FWTPct(3) = 12
            FWTRange(4) = 93250: FWTAmount(4) = 9328: FWTPct(4) = 22
            FWTRange(5) = 184950: FWTAmount(5) = 29502: FWTPct(5) = 24
            FWTRange(6) = 342050: FWTAmount(6) = 67206: FWTPct(6) = 32
            FWTRange(7) = 431050: FWTAmount(7) = 95686: FWTPct(7) = 35
            FWTRange(8) = 640500: FWTAmount(8) = 168993.5: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2021
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2021!", vbOKOnly + vbInformation

End Sub


Private Sub FWT2020Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 3800: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 13675: FWTAmount(3) = 987.5: FWTPct(3) = 12
            FWTRange(4) = 43925: FWTAmount(4) = 4617.5: FWTPct(4) = 22
            FWTRange(5) = 89325: FWTAmount(5) = 14605.5: FWTPct(5) = 24
            FWTRange(6) = 167100: FWTAmount(6) = 33271.5: FWTPct(6) = 32
            FWTRange(7) = 211150: FWTAmount(7) = 47367.5: FWTPct(7) = 35
            FWTRange(8) = 522200: FWTAmount(8) = 156235: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 11900: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 31650: FWTAmount(3) = 1975: FWTPct(3) = 12
            FWTRange(4) = 92150: FWTAmount(4) = 9235: FWTPct(4) = 22
            FWTRange(5) = 182950: FWTAmount(5) = 29211: FWTPct(5) = 24
            FWTRange(6) = 338500: FWTAmount(6) = 66543: FWTPct(6) = 32
            FWTRange(7) = 426600: FWTAmount(7) = 94735: FWTPct(7) = 35
            FWTRange(8) = 633950: FWTAmount(8) = 167307.5: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2020
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2020!", vbOKOnly + vbInformation

End Sub


Private Sub FWT2019Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 3800: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 13500: FWTAmount(3) = 970#: FWTPct(3) = 12
            FWTRange(4) = 43275: FWTAmount(4) = 4543#: FWTPct(4) = 22
            FWTRange(5) = 88000: FWTAmount(5) = 14382.5: FWTPct(5) = 24
            FWTRange(6) = 164525: FWTAmount(6) = 32748.5: FWTPct(6) = 32
            FWTRange(7) = 207900: FWTAmount(7) = 46628.5: FWTPct(7) = 35
            FWTRange(8) = 514100: FWTAmount(8) = 153798.5: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 11800: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 31200: FWTAmount(3) = 1940: FWTPct(3) = 12
            FWTRange(4) = 90750: FWTAmount(4) = 9086: FWTPct(4) = 22
            FWTRange(5) = 180200: FWTAmount(5) = 28765: FWTPct(5) = 24
            FWTRange(6) = 333250: FWTAmount(6) = 65497: FWTPct(6) = 32
            FWTRange(7) = 420000: FWTAmount(7) = 93257: FWTPct(7) = 35
            FWTRange(8) = 624150: FWTAmount(8) = 164709.5: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2019
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2019!", vbOKOnly + vbInformation

End Sub


Private Sub FWT2018Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 3700: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 13225: FWTAmount(3) = 952.5: FWTPct(3) = 12
            FWTRange(4) = 42400: FWTAmount(4) = 4453.5: FWTPct(4) = 22
            FWTRange(5) = 86200: FWTAmount(5) = 14089.5: FWTPct(5) = 24
            FWTRange(6) = 161200: FWTAmount(6) = 32089.5: FWTPct(6) = 32
            FWTRange(7) = 203700: FWTAmount(7) = 45689.5: FWTPct(7) = 35
            FWTRange(8) = 503700: FWTAmount(8) = 150689.5: FWTPct(8) = 37
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 11550: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 30600: FWTAmount(3) = 1905: FWTPct(3) = 12
            FWTRange(4) = 88950: FWTAmount(4) = 8907: FWTPct(4) = 22
            FWTRange(5) = 176550: FWTAmount(5) = 28179: FWTPct(5) = 24
            FWTRange(6) = 326550: FWTAmount(6) = 64179: FWTPct(6) = 32
            FWTRange(7) = 411550: FWTAmount(7) = 91379: FWTPct(7) = 35
            FWTRange(8) = 611550: FWTAmount(8) = 161379: FWTPct(8) = 37
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2018
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2018!", vbOKOnly + vbInformation

End Sub

Private Sub FWT2017Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2300: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 11625: FWTAmount(3) = 932.5: FWTPct(3) = 15
            FWTRange(4) = 40250: FWTAmount(4) = 5226.25: FWTPct(4) = 25
            FWTRange(5) = 94200: FWTAmount(5) = 18713.75: FWTPct(5) = 28
            FWTRange(6) = 193950: FWTAmount(6) = 46643.75: FWTPct(6) = 33
            FWTRange(7) = 419000: FWTAmount(7) = 120910.25: FWTPct(7) = 35
            FWTRange(8) = 420700: FWTAmount(8) = 121505.25: FWTPct(8) = 39.6
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 8650: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 27300: FWTAmount(3) = 1865: FWTPct(3) = 15
            FWTRange(4) = 84550: FWTAmount(4) = 10452.5: FWTPct(4) = 25
            FWTRange(5) = 161750: FWTAmount(5) = 29752.5: FWTPct(5) = 28
            FWTRange(6) = 242000: FWTAmount(6) = 52222.5: FWTPct(6) = 33
            FWTRange(7) = 425350: FWTAmount(7) = 112728: FWTPct(7) = 35
            FWTRange(8) = 479350: FWTAmount(8) = 131628: FWTPct(8) = 39.6
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2017
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2017!", vbOKOnly + vbInformation

End Sub


Private Sub FWT2016Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2250: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 11525: FWTAmount(3) = 927.5: FWTPct(3) = 15
            FWTRange(4) = 39900: FWTAmount(4) = 5183.75: FWTPct(4) = 25
            FWTRange(5) = 93400: FWTAmount(5) = 18558.75: FWTPct(5) = 28
            FWTRange(6) = 192400: FWTAmount(6) = 46278.75: FWTPct(6) = 33
            FWTRange(7) = 415600: FWTAmount(7) = 119934.75: FWTPct(7) = 35
            FWTRange(8) = 417300: FWTAmount(8) = 120529.75: FWTPct(8) = 39.6
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 8550: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 27100: FWTAmount(3) = 1855: FWTPct(3) = 15
            FWTRange(4) = 83850: FWTAmount(4) = 10367.5: FWTPct(4) = 25
            FWTRange(5) = 160450: FWTAmount(5) = 29517.5: FWTPct(5) = 28
            FWTRange(6) = 240000: FWTAmount(6) = 51791.5: FWTPct(6) = 33
            FWTRange(7) = 421900: FWTAmount(7) = 111818.5: FWTPct(7) = 35
            FWTRange(8) = 475500: FWTAmount(8) = 130578.5: FWTPct(8) = 39.6
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2016
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

    MsgBox "Federal tax tables updated for 2016!", vbOKOnly + vbInformation

End Sub

Private Sub FWT2015Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2300: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 11525: FWTAmount(3) = 922.5: FWTPct(3) = 15
            FWTRange(4) = 39750: FWTAmount(4) = 5156.25: FWTPct(4) = 25
            FWTRange(5) = 93050: FWTAmount(5) = 18481.25: FWTPct(5) = 28
            FWTRange(6) = 191600: FWTAmount(6) = 46075.25: FWTPct(6) = 33
            FWTRange(7) = 413800: FWTAmount(7) = 119401.25: FWTPct(7) = 35
            FWTRange(8) = 415500: FWTAmount(8) = 119996.25: FWTPct(8) = 39.6
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 8600: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 27050: FWTAmount(3) = 1845: FWTPct(3) = 15
            FWTRange(4) = 83500: FWTAmount(4) = 10312.5: FWTPct(4) = 25
            FWTRange(5) = 159800: FWTAmount(5) = 29387.5: FWTPct(5) = 28
            FWTRange(6) = 239050: FWTAmount(6) = 51577.5: FWTPct(6) = 33
            FWTRange(7) = 420100: FWTAmount(7) = 111324: FWTPct(7) = 35
            FWTRange(8) = 473450: FWTAmount(8) = 129996.5: FWTPct(8) = 39.6
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2015
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub

Private Sub FWT2013Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2200: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 11125: FWTAmount(3) = 892.5: FWTPct(3) = 15
            FWTRange(4) = 38450: FWTAmount(4) = 4991.25: FWTPct(4) = 25
            FWTRange(5) = 90050: FWTAmount(5) = 17891.25: FWTPct(5) = 28
            FWTRange(6) = 185450: FWTAmount(6) = 44603.25: FWTPct(6) = 33
            FWTRange(7) = 400550: FWTAmount(7) = 115586.25: FWTPct(7) = 35
            FWTRange(8) = 402200: FWTAmount(8) = 116163.75: FWTPct(8) = 39.6
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 8300: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 26150: FWTAmount(3) = 1785: FWTPct(3) = 15
            FWTRange(4) = 80800: FWTAmount(4) = 9982.5: FWTPct(4) = 25
            FWTRange(5) = 154700: FWTAmount(5) = 28457.5: FWTPct(5) = 28
            FWTRange(6) = 231350: FWTAmount(6) = 49919.5: FWTPct(6) = 33
            FWTRange(7) = 406650: FWTAmount(7) = 107768.5: FWTPct(7) = 35
            FWTRange(8) = 458300: FWTAmount(8) = 125846#: FWTPct(8) = 39.6
        End If

        For Lvl = 1 To 8

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2013
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 8 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub

Private Sub FWT2012Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2150: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 10850: FWTAmount(3) = 870: FWTPct(3) = 15
            FWTRange(4) = 37500: FWTAmount(4) = 4867.5: FWTPct(4) = 25
            FWTRange(5) = 87800: FWTAmount(5) = 17442.5: FWTPct(5) = 28
            FWTRange(6) = 180800: FWTAmount(6) = 43482.5: FWTPct(6) = 33
            FWTRange(7) = 390500: FWTAmount(7) = 112683.5: FWTPct(7) = 35
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 8100: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 25500: FWTAmount(3) = 1740: FWTPct(3) = 15
            FWTRange(4) = 78800: FWTAmount(4) = 9735: FWTPct(4) = 25
            FWTRange(5) = 150800: FWTAmount(5) = 27735: FWTPct(5) = 28
            FWTRange(6) = 225550: FWTAmount(6) = 48665: FWTPct(6) = 33
            FWTRange(7) = 396450: FWTAmount(7) = 105062: FWTPct(7) = 35
        End If

        For Lvl = 1 To 7

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2012
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 7 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub


Private Sub FWT2011Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 2100: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 10600: FWTAmount(3) = 850: FWTPct(3) = 15
            FWTRange(4) = 36600: FWTAmount(4) = 4750: FWTPct(4) = 25
            FWTRange(5) = 85700: FWTAmount(5) = 17025: FWTPct(5) = 28
            FWTRange(6) = 176500: FWTAmount(6) = 42449: FWTPct(6) = 33
            FWTRange(7) = 381250: FWTAmount(7) = 110016.5: FWTPct(7) = 35
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 7900: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 24900: FWTAmount(3) = 1700: FWTPct(3) = 15
            FWTRange(4) = 76900: FWTAmount(4) = 9500: FWTPct(4) = 25
            FWTRange(5) = 147250: FWTAmount(5) = 27087.5: FWTPct(5) = 28
            FWTRange(6) = 220200: FWTAmount(6) = 47513.5: FWTPct(6) = 33
            FWTRange(7) = 387050: FWTAmount(7) = 102574: FWTPct(7) = 35
        End If

        For Lvl = 1 To 7

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2011
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 7 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub

Private Sub FWT2010Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 6050: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 10425: FWTAmount(3) = 437.5: FWTPct(3) = 15
            FWTRange(4) = 36050: FWTAmount(4) = 4281.25: FWTPct(4) = 25
            FWTRange(5) = 67700: FWTAmount(5) = 12193.75: FWTPct(5) = 27
            FWTRange(6) = 84450: FWTAmount(6) = 16716.25: FWTPct(6) = 30
            FWTRange(7) = 87700: FWTAmount(7) = 17691.25: FWTPct(7) = 28
            FWTRange(8) = 173900: FWTAmount(8) = 41827.25: FWTPct(8) = 33
            FWTRange(9) = 375700: FWTAmount(9) = 108421.25: FWTPct(9) = 35
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0
            FWTRange(2) = 13750: FWTAmount(2) = 0: FWTPct(2) = 10
            FWTRange(3) = 24500: FWTAmount(3) = 1075: FWTPct(3) = 15
            FWTRange(4) = 75750: FWTAmount(4) = 8762.5: FWTPct(4) = 25
            FWTRange(5) = 94050: FWTAmount(5) = 13337.5: FWTPct(5) = 27
            FWTRange(6) = 124050: FWTAmount(6) = 21437.5: FWTPct(6) = 25
            FWTRange(7) = 145050: FWTAmount(7) = 26687.5: FWTPct(7) = 28
            FWTRange(8) = 217000: FWTAmount(8) = 46833.5: FWTPct(8) = 33
            FWTRange(9) = 381400: FWTAmount(9) = 101085.5: FWTPct(9) = 35
        End If

        For Lvl = 1 To 9

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2010
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 0

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 9 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub
Private Sub SWTOH2013Update()

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.638
    FWTRange(2) = 5000.99: FWTAmount(2) = 31.9: FWTPct(2) = 1.276
    FWTRange(3) = 10000.99: FWTAmount(3) = 95.7: FWTPct(3) = 2.552
    FWTRange(4) = 15000.99: FWTAmount(4) = 223.3: FWTPct(4) = 3.19
    FWTRange(5) = 20000.99: FWTAmount(5) = 382.8: FWTPct(5) = 3.828
    FWTRange(6) = 40000.99: FWTAmount(6) = 1148.4: FWTPct(6) = 4.466
    FWTRange(7) = 80000.99: FWTAmount(7) = 2934.8: FWTPct(7) = 5.103
    FWTRange(8) = 100000.99: FWTAmount(8) = 3955.4: FWTPct(8) = 6.379

    For Lvl = 1 To 8

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2013
        PRFWTTable.TaxMonth = 1
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 8 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

End Sub
Private Sub SWTOH2013UpdateSep1()

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.581
    FWTRange(2) = 5000.99: FWTAmount(2) = 29.05: FWTPct(2) = 1.161
    FWTRange(3) = 10000.99: FWTAmount(3) = 87.1: FWTPct(3) = 2.322
    FWTRange(4) = 15000.99: FWTAmount(4) = 203.2: FWTPct(4) = 2.903
    FWTRange(5) = 20000.99: FWTAmount(5) = 348.35: FWTPct(5) = 3.483
    FWTRange(6) = 40000.99: FWTAmount(6) = 1044.95: FWTPct(6) = 4.064
    FWTRange(7) = 80000.99: FWTAmount(7) = 2670.55: FWTPct(7) = 4.644
    FWTRange(8) = 100000.99: FWTAmount(8) = 3599.35: FWTPct(8) = 5.805

    For Lvl = 1 To 8

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2013
        PRFWTTable.TaxMonth = 9
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 8 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

    MsgBox "Ohio SWT tables updated for Sep 2013!", vbOKOnly + vbInformation

End Sub
Private Sub SWTOH2014UpdateJul1()

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.574
    FWTRange(2) = 5000.99: FWTAmount(2) = 28.7: FWTPct(2) = 1.148
    FWTRange(3) = 10000.99: FWTAmount(3) = 86.1: FWTPct(3) = 2.297
    FWTRange(4) = 15000.99: FWTAmount(4) = 200.95: FWTPct(4) = 2.871
    FWTRange(5) = 20000.99: FWTAmount(5) = 344.5: FWTPct(5) = 3.445
    FWTRange(6) = 40000.99: FWTAmount(6) = 1033.5: FWTPct(6) = 4.019
    FWTRange(7) = 80000.99: FWTAmount(7) = 2641.1: FWTPct(7) = 4.593
    FWTRange(8) = 100000.99: FWTAmount(8) = 3559.7: FWTPct(8) = 5.741

    For Lvl = 1 To 8

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2014
        PRFWTTable.TaxMonth = 7
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 8 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

    MsgBox "Ohio SWT tables updated for July 2014!", vbOKOnly + vbInformation

End Sub
Private Sub SWTOH2024Update()

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.501
    FWTRange(2) = 5000.99: FWTAmount(2) = 25.05: FWTPct(2) = 1.001
    FWTRange(3) = 10000.99: FWTAmount(3) = 75.1: FWTPct(3) = 2.005
    FWTRange(4) = 15000.99: FWTAmount(4) = 175.35: FWTPct(4) = 2.505
    FWTRange(5) = 20000.99: FWTAmount(5) = 300.6: FWTPct(5) = 2.99
    FWTRange(6) = 100000.99: FWTAmount(6) = 2692.6: FWTPct(6) = 4.41

    For Lvl = 1 To 6

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2024
        PRFWTTable.TaxMonth = 1
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 8 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

    MsgBox "Ohio SWT tables updated for Jan 2024!", vbOKOnly + vbInformation

End Sub

Private Sub SWTOH2019Update()

    Dim Multiplier As Double
    Multiplier = 1.075

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.5
    FWTRange(2) = 5000.99: FWTAmount(2) = 25: FWTPct(2) = 1
    FWTRange(3) = 10000.99: FWTAmount(3) = 75: FWTPct(3) = 2
    FWTRange(4) = 15000.99: FWTAmount(4) = 175: FWTPct(4) = 2.5
    FWTRange(5) = 20000.99: FWTAmount(5) = 300: FWTPct(5) = 3
    FWTRange(6) = 40000.99: FWTAmount(6) = 900: FWTPct(6) = 3.5
    FWTRange(7) = 80000.99: FWTAmount(7) = 2300: FWTPct(7) = 4
    FWTRange(8) = 100000.99: FWTAmount(8) = 3100: FWTPct(8) = 5

    For Lvl = 1 To 8

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2019
        PRFWTTable.TaxMonth = 1
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 8 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

    MsgBox "Ohio SWT tables updated for Jan 2019!", vbOKOnly + vbInformation

End Sub


Private Sub SWTOH2015UpdateAug1()

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.556
    FWTRange(2) = 5000.99: FWTAmount(2) = 27.8: FWTPct(2) = 1.112
    FWTRange(3) = 10000.99: FWTAmount(3) = 83.4: FWTPct(3) = 2.226
    FWTRange(4) = 15000.99: FWTAmount(4) = 194.7: FWTPct(4) = 2.782
    FWTRange(5) = 20000.99: FWTAmount(5) = 333.8: FWTPct(5) = 3.338
    FWTRange(6) = 40000.99: FWTAmount(6) = 1001.4: FWTPct(6) = 3.894
    FWTRange(7) = 80000.99: FWTAmount(7) = 2559#: FWTPct(7) = 4.451
    FWTRange(8) = 100000.99: FWTAmount(8) = 3449.2: FWTPct(8) = 5.563

    For Lvl = 1 To 8

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2015
        PRFWTTable.TaxMonth = 8
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 8 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

    MsgBox "Ohio SWT tables updated for August 2015!", vbOKOnly + vbInformation

End Sub

Private Sub SWTOH2010Update()

    FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 0.618
    FWTRange(2) = 5000.99: FWTAmount(2) = 30.9: FWTPct(2) = 1.236
    FWTRange(3) = 10000.99: FWTAmount(3) = 92.7: FWTPct(3) = 2.473
    FWTRange(4) = 15000.99: FWTAmount(4) = 216.35: FWTPct(4) = 3.091
    FWTRange(5) = 20000.99: FWTAmount(5) = 370.9: FWTPct(5) = 3.708
    FWTRange(6) = 40000.99: FWTAmount(6) = 1112.5: FWTPct(6) = 4.327
    FWTRange(7) = 80000.99: FWTAmount(7) = 2843.3: FWTPct(7) = 4.945
    FWTRange(8) = 100000.99: FWTAmount(8) = 3832.3: FWTPct(8) = 5.741
    FWTRange(9) = 200000.99: FWTAmount(9) = 9573.3: FWTPct(9) = 6.24

    For Lvl = 1 To 9

        PRFWTTable.Clear
        PRFWTTable.TaxYear = 2010
        PRFWTTable.TaxMonth = 1
        PRFWTTable.StateID = 36

        If Lvl = 1 Then
            PRFWTTable.LowAmount = 0
            PRFWTTable.ExcessBase = 0
        Else
            PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
            PRFWTTable.ExcessBase = Int(FWTRange(Lvl))
        End If

        If Lvl = 9 Then
            PRFWTTable.HiAmount = 99999999.99
        Else
            PRFWTTable.HiAmount = FWTRange(Lvl + 1)
        End If

        PRFWTTable.Amount = FWTAmount(Lvl)
        PRFWTTable.Percent = FWTPct(Lvl)
        PRFWTTable.Save (Equate.RecAdd)

    Next Lvl

End Sub

Private Sub SWTMD2010Update()

    For SnglMarr = 1 To 2     ' 1 = single / 2 = married

        If SnglMarr = 1 Then
            ' FWT SINGLE
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 6
            FWTRange(2) = 150000: FWTAmount(2) = 9000: FWTPct(2) = 6.25
            FWTRange(3) = 300000: FWTAmount(3) = 18375: FWTPct(3) = 6.5
            FWTRange(4) = 500000: FWTAmount(4) = 31375: FWTPct(4) = 6.75
            FWTRange(5) = 1000000: FWTAmount(5) = 65125: FWTPct(5) = 7.5
        Else
            ' FWT MARRIED
            FWTRange(1) = 0: FWTAmount(1) = 0: FWTPct(1) = 6
            FWTRange(2) = 200000: FWTAmount(2) = 12000: FWTPct(2) = 6.25
            FWTRange(3) = 350000: FWTAmount(3) = 21375: FWTPct(3) = 6.5
            FWTRange(4) = 500000: FWTAmount(4) = 31125: FWTPct(4) = 6.75
            FWTRange(5) = 1000000: FWTAmount(5) = 64875: FWTPct(5) = 7.5
        End If

        For Lvl = 1 To 5

            PRFWTTable.Clear
            PRFWTTable.TaxYear = 2010
            PRFWTTable.TaxMonth = 1
            PRFWTTable.StateID = 21

            If SnglMarr = 1 Then
                PRFWTTable.msSingle = 1
                PRFWTTable.msMarried = 0
            Else
                PRFWTTable.msSingle = 0
                PRFWTTable.msMarried = 1
            End If

            If Lvl = 1 Then
                PRFWTTable.LowAmount = 0
                PRFWTTable.ExcessBase = 0
            Else
                PRFWTTable.LowAmount = FWTRange(Lvl) + 0.01
                PRFWTTable.ExcessBase = FWTRange(Lvl)
            End If

            If Lvl = 5 Then
                PRFWTTable.HiAmount = 99999999.99
            Else
                PRFWTTable.HiAmount = FWTRange(Lvl + 1)
            End If

            PRFWTTable.Amount = FWTAmount(Lvl)
            PRFWTTable.Percent = FWTPct(Lvl)
            PRFWTTable.Save (Equate.RecAdd)

        Next Lvl

    Next SnglMarr

End Sub

