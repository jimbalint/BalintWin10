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

    ' 12/30/09 - 2010 tax update
    ' 12/31/10 - 2011 tax update
    ' 02/07/12 - 2012 tax update
    ' 01/13/13 - 2013 tax update
    If GLSys = False Then
        boo = AddField("PRHist", "MedAddAmt", "Currency", adoConn)
    End If
    
    If GLSys = True Then

        ' --- 2016 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
                    "AND Year = 2016"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
            PRGlobal.Year = 2016
            PRGlobal.Description = "FWT ALLOW"
            PRGlobal.Amount = 4050
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "FWT Allowance for 2016 updated to: $4,050", vbInformation
        End If
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2016"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2016
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 118500#
            PRGlobal.Save (Equate.RecAdd)
            MsgBox "SS Max for 2016 updated to: $118,500", vbInformation
        End If

        ' --- 2015 ---------------------------------------------------------------
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFWTAllow & " " & _
                    "AND Year = 2015"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeFWTAllow
            PRGlobal.Year = 2015
            PRGlobal.Description = "FWT ALLOW"
            PRGlobal.Amount = 4000
            PRGlobal.Save (Equate.RecAdd)
        End If

        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeSSMax & " " & _
                    "AND Year = 2015"
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeSSMax
            PRGlobal.Year = 2015
            PRGlobal.Description = "SS MAX"
            PRGlobal.Amount = 118500#
            PRGlobal.Save (Equate.RecAdd)
        End If

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
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2016 AND StateID = 0"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2016Update
        
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2015 AND StateID = 0"
        If PRFWTTable.GetBySQL(SQLString) = False Then FWT2015Update
        
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
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2010 AND StateID = 36"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2010Update

        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2013 AND StateID = 36"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2013Update
    
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2013 AND StateID = 36 AND TaxMonth = 9"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2013UpdateSep1
    
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2014 AND StateID = 36 AND TaxMonth = 7"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2014UpdateJul1
    
        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = 2015 AND StateID = 36 AND TaxMonth = 8"
        If PRFWTTable.GetBySQL(SQLString) = False Then SWTOH2015UpdateAug1
    
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
Dim FString As String
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
        
        FString = "ALTER TABLE " & TableName & _
                  " ADD COLUMN [" & ColumnName & "]" & _
                  " " & ColumnType
        adoConn.Execute FString
        
        If Err.Number = 0 Then
            AddField = 1
            Exit Do
        Else
            If InStr(1, LCase(Err.Description), "could not lock", vbTextCompare) Then
                MsgResponse = MsgBox("Database update not complete" & vbCr & _
                              "ALL other users must exit to proceed!", vbRetryCancel + vbExclamation)
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

