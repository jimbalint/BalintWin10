Attribute VB_Name = "modDBCreate"
Option Explicit

Public Sub CompanyCreate()

    SQLString = "CREATE TABLE PRCompany ( " & _
                        "[CompanyID] Counter, CONSTRAINT prcIDKey PRIMARY KEY ([CompanyID]) ) "
                        
    cnDes.Execute SQLString
                        
    AddField "PRCompany", "Name", "Char (255)", cnDes
    AddField "PRCompany", "Address1", "Char (255)", cnDes
    AddField "PRCompany", "Address2", "Char (255)", cnDes
    AddField "PRCompany", "City", "Char (255)", cnDes
    AddField "PRCompany", "AddrStateID", "Long", cnDes
    AddField "PRCompany", "ZipCode", "Long", cnDes
    AddField "PRCompany", "PhoneNumber", "Char (255)", cnDes
    AddField "PRCompany", "StateID", "Char (25)", cnDes
    AddField "PRCompany", "StateUnempPct", "Double", cnDes
    AddField "PRCompany", "StateUnempID", "Char (255)", cnDes
    AddField "PRCompany", "FederalID", "Char (25)", cnDes
    AddField "PRCompany", "FederalUnempPct", "Double", cnDes
    AddField "PRCompany", "DfltPaysPerYear", "Long", cnDes
    AddField "PRCompany", "DfltStateID", "Long", cnDes
    AddField "PRCompany", "DfltMinWage", "Currency", cnDes
    AddField "PRCompany", "DfltOTRate", "Currency", cnDes
    AddField "PRCompany", "DfltRegHrs", "Currency", cnDes
    AddField "PRCompany", "FileName", "Char (255)", cnDes
    
    AddField "PRCompany", "GLAcctSS", "Long", cnDes
    AddField "PRCompany", "GLAcctMED", "Long", cnDes
    AddField "PRCompany", "GLAcctFWT", "Long", cnDes
    AddField "PRCompany", "GLAcctSWT", "Long", cnDes
    AddField "PRCompany", "GLAcctCWT", "Long", cnDes
    AddField "PRCompany", "GLAcctGross", "Long", cnDes
    AddField "PRCompany", "GLAcctNet", "Long", cnDes

    AddField "PRCompany", "BankName", "String", cnDes
    AddField "PRCompany", "BankABA", "String", cnDes
    AddField "PRCompany", "BankAccount", "String", cnDes
    AddField "PRCompany", "BankAddr1", "String", cnDes
    AddField "PRCompany", "BankAddr2", "String", cnDes
    AddField "PRCompany", "BankFraction", "String", cnDes

    AddField "PRCompany", "LastCheckNum", "Long", cnDes
    AddField "PRCompany", "DfltCityID", "Long", cnDes
    AddField "PRCompany", "CheckDays", "Long", cnDes

    AddField "PRCompany", "WkcPolicyNum", "String", cnDes

    AddField "PRCompany", "GLCompanyID", "Long", cnDes
    AddField "PRCompany", "DfltSortOrder", "Byte", cnDes

    ' AddField "PRCompany", "Comment1", "Char (100)", cnDes
    ' AddField "PRCompany", "Comment2", "Char (100)", cnDes

    AddField "PRCompany", "DirDepBalanced", "Byte", cnDes
    AddField "PRCompany", "DirDepUseAltID", "Byte", cnDes
    AddField "PRCompany", "DirDepAltID", "Long", cnDes
    
    ' include a "1" before the Fed ID???
    AddField "PRCompany", "DirDepID1", "Byte", cnDes
    
End Sub

Public Sub EmployeeCreate()

    
    SQLString = "CREATE TABLE PREmployee ( " & _
                        "[EmployeeID]       Counter, CONSTRAINT empIDKey PRIMARY KEY ([EmployeeID]), " & _
                        "[EmployeeNumber]   Long, CONSTRAINT empNumberKey UNIQUE ([EmployeeNumber]) ) "
    
    cn.Execute SQLString
    
    AddField "PREmployee", "LastName", "Char (255)", cn
    AddField "PREmployee", "FirstName", "Char (255)", cn
    AddField "PREmployee", "MidInit", "Char (2)", cn
    AddField "PREmployee", "AltName", "Char (255)", cn
    AddField "PREmployee", "UseAltName", "Byte", cn
    AddField "PREmployee", "Address1", "Char (255)", cn
    AddField "PREmployee", "Address2", "Char (255)", cn
    AddField "PREmployee", "City", "Char (255)", cn
    AddField "PREmployee", "State", "Char (2)", cn
    AddField "PREmployee", "ZipCode", "Long", cn
    
    If NewADO Then
        AddField "PREmployee", "SSN", "Char (255)", cn
    Else
        AddField "PREmployee", "SSN", "Long", cn
    End If
    
    AddField "PREmployee", "DepartmentID", "Long", cn
    AddField "PREmployee", "SalaryAmount", "Currency", cn
    AddField "PREmployee", "HourlyAmount", "Currency", cn
    AddField "PREmployee", "Inactive", "Byte", cn
    AddField "PREmployee", "Salaried", "Byte", cn
    AddField "PREmployee", "PaysPerYear", "Byte", cn
    AddField "PREmployee", "NoSSTax", "Byte", cn
    AddField "PREmployee", "NoMedTax", "Byte", cn
    AddField "PREmployee", "NoFedTax", "Byte", cn
    AddField "PREmployee", "NoStateTax", "Byte", cn
    AddField "PREmployee", "NoCityTax", "Byte", cn
    AddField "PREmployee", "NoFedUnemp", "Byte", cn
    AddField "PREmployee", "NoStateUnemp", "Byte", cn
    
    AddField "PREmployee", "FWTMarried", "Byte", cn
    AddField "PREmployee", "FWTBasis", "Byte", cn
    AddField "PREmployee", "FWTAmount", "Currency", cn
    AddField "PREmployee", "FWTExtraBasis", "Byte", cn
    AddField "PREmployee", "FWTExtraAmount", "Currency", cn
    
    AddField "PREmployee", "SWTMarried", "Byte", cn
    AddField "PREmployee", "SWTBasis", "Byte", cn
    AddField "PREmployee", "SWTAmount", "Currency", cn
    AddField "PREmployee", "SWTExtraBasis", "Byte", cn
    AddField "PREmployee", "SWTExtraAmount", "Currency", cn
    
    AddField "PREmployee", "DefaultCityID", "Long", cn
    AddField "PREmployee", "DefaultJobID", "Long", cn
    AddField "PREmployee", "CourtesyCityID", "Long", cn
    AddField "PREmployee", "CourtesyAdd", "Byte", cn

    AddField "PREmployee", "x1099Employee", "Byte", cn
    AddField "PREmployee", "Statutory", "Byte", cn
    AddField "PREmployee", "EICType", "Byte", cn
    AddField "PREmployee", "WkcUseDept", "Byte", cn
    AddField "PREmployee", "WkcCat", "Long", cn

    AddField "PREmployee", "DateLastPaid", "DateTime", cn
    AddField "PREmployee", "DateHired", "DateTime", cn
    AddField "PREmployee", "DateLastRaise", "DateTime", cn
    AddField "PREmployee", "DateLastReview", "DateTime", cn
    AddField "PREmployee", "DateLastLayoff", "DateTime", cn
    AddField "PREmployee", "DateLastRecall", "DateTime", cn
    AddField "PREmployee", "DateTerminated", "DateTime", cn
    AddField "PREmployee", "DateOfBirth", "DateTime", cn

    AddField "PREmployee", "TermReason", "Long", cn
    AddField "PREmployee", "Sex", "Char (1)", cn
    AddField "PREmployee", "RaceCode", "Long", cn
    AddField "PREmployee", "MaritalStatus", "Char (1)", cn
    
    AddField "PREmployee", "EducationLevel", "Long", cn
    AddField "PREmployee", "ShiftCode", "Long", cn
    
    AddField "PREmployee", "WorkCompNum", "Long", cn
    
    ' 2020-04-16
    ' AddField "PREmployee", "CheckComment", "Char (255)", cn
    AddField "PREmployee", "CheckComment", "LongText", cn

End Sub

Public Sub DepartmentCreate()

    
    SQLString = "CREATE TABLE PRDepartment ( " & _
                        "[DepartmentID]            Counter, CONSTRAINT dptIDKey PRIMARY KEY ([DepartmentID]), " & _
                        "[DepartmentNumber]        Long, CONSTRAINT dptNumberKey UNIQUE ([DepartmentNumber]) ) "
    
    cn.Execute SQLString
    
    AddField "PRDepartment", "Name", "Char (255)", cn
    AddField "PRDepartment", "WkcCat", "Long", cn

End Sub

Public Sub ItemCreate()

    SQLString = "CREATE TABLE PRItem ( " & _
                        "[ItemID]            Counter, CONSTRAINT itemIDKey PRIMARY KEY ([ItemID]) )"
    
    cn.Execute SQLString
    
    AddField "PRItem", "EmployeeID", "Long", cn
    AddField "PRItem", "Title", "Char (255)", cn
    AddField "PRItem", "Abbreviation", "Char (255)", cn
    AddField "PRItem", "ItemType", "Byte", cn
    AddField "PRItem", "Active", "Byte", cn
    AddField "PRItem", "NoSSTax", "Byte", cn
    AddField "PRItem", "NoMedTax", "Byte", cn
    AddField "PRItem", "NoFWTTax", "Byte", cn
    AddField "PRItem", "NoSWTTax", "Byte", cn
    AddField "PRItem", "NoCWTTax", "Byte", cn
    AddField "PRItem", "NoSUNTax", "Byte", cn
    AddField "PRItem", "NoFUNTax", "Byte", cn
    AddField "PRItem", "GLAccount", "Long", cn
    AddField "PRItem", "Basis", "Long", cn
    AddField "PRItem", "MatchPct", "Double", cn
    AddField "PRItem", "MaxPct", "Double", cn
    AddField "PRItem", "MaxAmount", "Currency", cn
    AddField "PRItem", "AmtPct", "Currency", cn
    AddField "PRItem", "Tips", "Byte", cn
    AddField "PRItem", "NotInNet", "Byte", cn
    AddField "PRItem", "DirDepType", "Byte", cn
    AddField "PRItem", "DirDepBank", "Char (255)", cn
    AddField "PRItem", "DirDepABA", "Char (255)", cn
    AddField "PRItem", "DirDepAccount", "Char (255)", cn
    AddField "PRItem", "DirDepBasis", "Byte", cn
    AddField "PRItem", "DirDepAmtPct", "Currency", cn
    AddField "PRItem", "W2Box12Code", "Char (10)", cn
    AddField "PRItem", "W2Box14Code", "Char (10)", cn
    AddField "PRItem", "Pension", "Byte", cn
    AddField "PRItem", "SickPay", "Byte", cn
    AddField "PRItem", "SDNumber", "Byte", cn
    AddField "PRItem", "EmployerItemID", "Long", cn
    AddField "PRItem", "UseEmployer", "Byte", cn
    AddField "PRItem", "Escrow", "Byte", cn
    AddField "PRItem", "Comment", "Char (50)", cn
    AddField "PRItem", "DirDepRpt", "Byte", cn
    AddField "PRItem", "RateDifference", "Byte", cn
    AddField "PRItem", "PWFringe", "Byte", cn
    AddField "PRItem", "CityID", "Long", cn

End Sub
Public Sub CityCreate()

    SQLString = "CREATE TABLE PRCity ( " & _
                        "[CityID]            Counter, CONSTRAINT itemIDKey PRIMARY KEY ([CityID]) )"
    
    cnDes.Execute SQLString
    
    AddField "PRCity", "CityNumber", "Long", cnDes
    AddField "PRCity", "CityName", "Char (255)", cnDes
    AddField "PRCity", "ShortName", "Char (20)", cnDes
    AddField "PRCity", "CityState", "Char (2)", cnDes
    AddField "PRCity", "CityRate", "Double", cnDes
    AddField "PRCity", "CityRecipRate", "Double", cnDes
    AddField "PRCity", "StateID", "Long", cnDes
    AddField "PRCity", "CountyID", "Long", cnDes
    
End Sub

Public Sub FFColumnCreate()

    SQLString = "CREATE TABLE GLFFColumn ( " & _
                        "[FFColumnID] Counter, CONSTRAINT colIDKey PRIMARY KEY ([FFColumnID]) )"
    
    cnDes.Execute SQLString
    
    AddField "GLFFColumn", "ColNum", "Byte", cnDes
    AddField "GLFFColumn", "Description", "Char (30)", cnDes
    AddField "GLFFColumn", "ColType", "Byte", cnDes
    AddField "GLFFColumn", "FiscalYear", "Long", cnDes
    AddField "GLFFColumn", "StartNum", "Byte", cnDes
    AddField "GLFFColumn", "EndNum", "Byte", cnDes
    AddField "GLFFColumn", "Budget", "Byte", cnDes
    AddField "GLFFColumn", "PrintTab", "Byte", cnDes
    AddField "GLFFColumn", "NonPrint", "Byte", cnDes
    AddField "GLFFColumn", "GlobalID", "Long", cnDes
    
End Sub


Public Sub StateCreate()

    SQLString = "CREATE TABLE PRState ( " & _
                        "[StateID]            Counter, CONSTRAINT itemIDKey PRIMARY KEY ([StateID]) )"
    
    cnDes.Execute SQLString
    
    AddField "PRState", "StateName", "Char (255)", cnDes
    AddField "PRState", "StateAbbrev", "Char (2)", cnDes
    AddField "PRState", "UnEmpMax", "Currency", cnDes
    
End Sub

Public Sub W2BoxCreate()

    SQLString = "CREATE TABLE PRW2Box ( " & _
                        "[W2BoxID]            Counter, CONSTRAINT itemIDKey PRIMARY KEY ([W2BoxID]) )"
    
    cnDes.Execute SQLString
    
    AddField "PRW2Box", "W2BoxNumber", "Long", cnDes
    AddField "PRW2Box", "W2BoxCode", "Char (20)", cnDes
    AddField "PRW2Box", "W2BoxDescription", "Char (255)", cnDes
    
End Sub

Public Sub HistCreate()

    SQLString = "CREATE TABLE PRHist ( " & _
                        "[HistID]            Counter, CONSTRAINT itemIDKey PRIMARY KEY ([HistID]) )"
    
    cn.Execute SQLString
    
    AddField "PRHist", "EmployeeID", "Long", cn
    AddField "PRHist", "BatchID", "Long", cn
    AddField "PRHist", "CheckNumber", "Long", cn
    AddField "PRHist", "StateID", "Long", cn
    
    AddField "PRHist", "YearMonth", "Long", cn
    AddField "PRHist", "CheckDate", "DateTime", cn
    AddField "PRHist", "PEDate", "DateTime", cn
    
    AddField "PRHist", "DepartmentID", "Long", cn
    
    AddField "PRHist", "RegHours", "Single", cn
    AddField "PRHist", "RegRate", "Currency", cn
    AddField "PRHist", "RegAmount", "Currency", cn
    
    AddField "PRHist", "OTHours", "Single", cn
    AddField "PRHist", "OTRate", "Currency", cn
    AddField "PRHist", "OTAmount", "Currency", cn
    
    AddField "PRHist", "OEHours", "Single", cn
    AddField "PRHist", "OERate", "Currency", cn
    AddField "PRHist", "OEAmount", "Currency", cn
    
    AddField "PRHist", "SSWageBase", "Currency", cn
    AddField "PRHist", "SSWage", "Currency", cn
    AddField "PRHist", "SSTax", "Currency", cn
    AddField "PRHist", "ManualSSTax", "Byte", cn
    
    AddField "PRHist", "MedWage", "Currency", cn
    AddField "PRHist", "MedTax", "Currency", cn
    AddField "PRHist", "ManualMedTax", "Byte", cn
    
    AddField "PRHist", "FWTWage", "Currency", cn
    AddField "PRHist", "FWTTax", "Currency", cn
    AddField "PRHist", "ManualFWTTax", "Byte", cn
    
    AddField "PRHist", "SWTWage", "Currency", cn
    AddField "PRHist", "SWTTax", "Currency", cn
    
    AddField "PRHist", "CWTWage", "Currency", cn
    AddField "PRHist", "CWTTax", "Currency", cn
    
    AddField "PRHist", "Deductions", "Currency", cn
    AddField "PRHist", "DirectDeposit", "Currency", cn
    
    AddField "PRHist", "Gross", "Currency", cn
    AddField "PRHist", "Net", "Currency", cn

    AddField "PRHist", "FUNWageBase", "Currency", cn
    AddField "PRHist", "FUNWage", "Currency", cn
    AddField "PRHist", "SUNWageBase", "Currency", cn
    AddField "PRHist", "SUNWage", "Currency", cn
    AddField "PRHist", "GLUpdate", "Byte", cn
    AddField "PRHist", "WkcAmount", "Currency", cn
    AddField "PRHist", "NotInNetAmount", "Currency", cn
    AddField "PRHist", "SDTax", "Currency", cn
    AddField "PRHist", "QBUpdateFlag", "Byte", cn

End Sub

Public Sub DistCreate()

    SQLString = "CREATE TABLE PRDist ( " & _
                        "[DistID]            Counter, CONSTRAINT distIDKey PRIMARY KEY ([DistID]) )"
    
    cn.Execute SQLString
    
    AddField "PRDist", "EmployeeID", "Long", cn
    AddField "PRDist", "BatchID", "Long", cn
    AddField "PRDist", "HistID", "Long", cn
    AddField "PRDist", "StateID", "Long", cn
    AddField "PRDist", "CityID", "Long", cn
    AddField "PRDist", "CourtesyCityID", "Long", cn
    AddField "PRDist", "JobID", "Long", cn
    AddField "PRDist", "CustomerID", "Long", cn
    AddField "PRDist", "DepartmentID", "Long", cn
    
    AddField "PRDist", "YearMonth", "Long", cn
    AddField "PRDist", "CheckDate", "DateTime", cn
    AddField "PRDist", "PEDate", "DateTime", cn
    
    AddField "PRDist", "DistType", "Byte", cn
    AddField "PRDist", "ItemID", "Long", cn
    AddField "PRDist", "EmployerItemID", "Long", cn
    AddField "PRDist", "ItemType", "Byte", cn
    AddField "PRDist", "Hours", "Single", cn
    AddField "PRDist", "Rate", "Currency", cn
    AddField "PRDist", "Amount", "Currency", cn
    AddField "PRDist", "ManualAmount", "Byte", cn
    
    AddField "PRDist", "BillingRate", "Currency", cn
    
    AddField "PRDist", "GrossWage", "Currency", cn
    
    AddField "PRDist", "StateWage", "Currency", cn
    AddField "PRDist", "StateTax", "Currency", cn
    AddField "PRDist", "ManualStateTax", "Byte", cn

    AddField "PRDist", "CityWage", "Currency", cn
    AddField "PRDist", "CityTax", "Currency", cn
    AddField "PRDist", "CourtesyCityTax", "Currency", cn
    AddField "PRDist", "ManualCityTax", "Byte", cn
    AddField "PRDist", "ManualCourtesyCityTax", "Byte", cn
    
    AddField "PRDist", "SUNWage", "Currency", cn
    
    AddField "PRDist", "HistFlag", "Byte", cn
    AddField "PRDist", "NotInNet", "Byte", cn
    AddField "PRDist", "SDTax", "Currency", cn

    AddField "PRDist", "QBInvoiceID", "Char (50)", cn

End Sub

Public Sub ItemHistCreate()

    SQLString = "CREATE TABLE PRItemHist ( " & _
                        "[ItemHistID]            Counter, CONSTRAINT itemIDKey PRIMARY KEY ([ItemHistID]) )"
    
    cn.Execute SQLString
    
    AddField "PRItemHist", "EmployeeID", "Long", cn
    AddField "PRItemHist", "HistID", "Long", cn
    AddField "PRItemHist", "BatchID", "Long", cn
    AddField "PRItemHist", "ItemID", "Long", cn
    AddField "PRItemHist", "DepartmentID", "Long", cn
    AddField "PRItemHist", "EmployerItemID", "Long", cn
    AddField "PRItemHist", "ItemType", "Byte", cn
    
    AddField "PRItemHist", "YearMonth", "Long", cn
    AddField "PRItemHist", "CheckDate", "DateTime", cn
    AddField "PRItemHist", "PEDate", "DateTime", cn
    
    AddField "PRItemHist", "Hours", "Single", cn
    AddField "PRItemHist", "Rate", "Currency", cn
    AddField "PRItemHist", "Amount", "Currency", cn
    AddField "PRItemHist", "ManualAmount", "Byte", cn
    AddField "PRItemHist", "Percent", "Currency", cn
    AddField "PRItemHist", "WageBase", "Currency", cn
    
    ' wage excluded from basis for deduct by percent (401k match purposes)
    AddField "PRItemHist", "WageExcluded", "Currency", cn
    
End Sub

Public Sub GLUpdCreate()

    SQLString = "CREATE TABLE PRGLUpd ( " & _
                        "[GLUpdID]            Counter, CONSTRAINT GLUpdIDKey PRIMARY KEY ([GLUpdID]) )"
    
    cn.Execute SQLString
    
    AddField "PRGLUpd", "GLType", "Byte", cn
    AddField "PRGLUpd", "RelatedID", "Long", cn
    AddField "PRGLUpd", "GLItemType", "Byte", cn
    AddField "PRGLUpd", "ItemID", "Long", cn
    AddField "PRGLUpd", "GLAccountNum", "Long", cn
    AddField "PRGLUpd", "Title", "Char(30)", cn

End Sub



Public Sub FWTCreate()

    SQLString = "CREATE TABLE PRFWTTable ( " & _
                        "[FWTTableID]            Counter, CONSTRAINT FWTTableIDKey PRIMARY KEY ([FWTTableID]) )"
    
    cnDes.Execute SQLString
    
    AddField "PRFWTTable", "StateID", "Long", cnDes
    AddField "PRFWTTable", "TaxYear", "Long", cnDes
    AddField "PRFWTTable", "TaxMonth", "Byte", cnDes
    
    AddField "PRFWTTable", "msMarried", "Byte", cnDes
    AddField "PRFWTTable", "msSingle", "Byte", cnDes
    AddField "PRFWTTable", "LowAmount", "Currency", cnDes
    AddField "PRFWTTable", "HiAmount", "Currency", cnDes
    AddField "PRFWTTable", "Amount", "Currency", cnDes
    AddField "PRFWTTable", "Percent", "Double", cnDes
    AddField "PRFWTTable", "ExcessBase", "Currency", cnDes

End Sub

Public Sub AdjustCreate()

    SQLString = "CREATE TABLE PRAdjust ( " & _
                        "[AdjustID]            Counter, CONSTRAINT PRAdjustIDKey PRIMARY KEY ([AdjustID]) )"
    
    cn.Execute SQLString
    
    AddField "PRAdjust", "EmployeeID", "Long", cn
    AddField "PRAdjust", "AdjDate", "DateTime", cn
    AddField "PRAdjust", "AdjType", "Byte", cn
    AddField "PRAdjust", "AdjAmount", "Currency", cn
    AddField "PRAdjust", "AdjHours", "Single", cn

End Sub

Public Sub PRBatchCreate()

    SQLString = "CREATE TABLE PRBatch ( " & _
                        "[BatchID]            Counter, CONSTRAINT PRBatchIDKey PRIMARY KEY ([BatchID]) )"
    
    cn.Execute SQLString
    
    AddField "PRBatch", "UserID", "Long", cn
    AddField "PRBatch", "CreateDate", "DateTime", cn
    AddField "PRBatch", "PEDate", "DateTime", cn
    AddField "PRBatch", "CheckDate", "DateTime", cn
    AddField "PRBatch", "RecCount", "Long", cn
    AddField "PRBatch", "YearMonth", "Long", cn
    AddField "PRBatch", "JobDist", "Byte", cn

End Sub

Public Sub EEListsCreate()

    SQLString = "CREATE TABLE PREELists ( " & _
                        "[EEListsID]            Counter, CONSTRAINT PREEListsIDKey PRIMARY KEY ([EEListsID]) )"
    
    cn.Execute SQLString
    
    AddField "PREELists", "EmployeeID", "Long", cn
    AddField "PREELists", "EEListsType", "Byte", cn
    AddField "PREELists", "EEListsString1", "String", cn
    AddField "PREELists", "EEListsString2", "String", cn
    AddField "PREELists", "EEListsAmount", "Currency", cn
    AddField "PREELists", "EEListsLong", "Long", cn

End Sub

Public Sub GlobalCreate()

    SQLString = "CREATE TABLE PRGlobal ( " & _
                        "[GlobalID]            Counter, CONSTRAINT PRGlobalIDKey PRIMARY KEY ([GlobalID]) )"
    
    cnDes.Execute SQLString
    
    AddField "PRGlobal", "TypeCode", "Byte", cnDes
    AddField "PRGlobal", "UserID", "Long", cnDes
    AddField "PRGlobal", "Description", "String", cnDes
    AddField "PRGlobal", "Amount", "Currency", cnDes
    AddField "PRGlobal", "Percent", "Double", cnDes
    AddField "PRGlobal", "Flag", "Byte", cnDes
    AddField "PRGlobal", "Year", "Long", cnDes
    AddField "PRGlobal", "Month", "Byte", cnDes
    AddField "PRGlobal", "Var1", "String", cnDes
    AddField "PRGlobal", "Var2", "String", cnDes
    AddField "PRGlobal", "Var3", "String", cnDes
    AddField "PRGlobal", "Var4", "String", cnDes
    AddField "PRGlobal", "Var5", "String", cnDes
    AddField "PRGlobal", "Var6", "String", cnDes
    AddField "PRGlobal", "Var7", "String", cnDes
    AddField "PRGlobal", "Var8", "String", cnDes
    AddField "PRGlobal", "Var9", "String", cnDes
    AddField "PRGlobal", "Var10", "String", cnDes
    
    AddField "PRGlobal", "Byte1", "Byte", cnDes
    AddField "PRGlobal", "Byte2", "Byte", cnDes
    AddField "PRGlobal", "Byte3", "Byte", cnDes
    AddField "PRGlobal", "Byte4", "Byte", cnDes
    AddField "PRGlobal", "Byte5", "Byte", cnDes
    AddField "PRGlobal", "Byte6", "Byte", cnDes
    AddField "PRGlobal", "Byte7", "Byte", cnDes
    AddField "PRGlobal", "Byte8", "Byte", cnDes
    AddField "PRGlobal", "Byte9", "Byte", cnDes
    AddField "PRGlobal", "Byte10", "Byte", cnDes

End Sub

Private Sub AddField(ByVal TableName As String, _
                     ByVal FieldName As String, _
                     ByVal FieldType As String, _
                     ByVal acn As ADODB.Connection)
                     
    SQLString = "ALTER TABLE " & TableName & _
              " ADD COLUMN [" & FieldName & "]   " & FieldType
    
    acn.Execute SQLString
                     
End Sub


Public Sub DropTable(ByVal TableName As String, _
                      ByVal adoCn As ADODB.Connection)

' *** Drop a table if it exists ***

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim TableFlag As Boolean
Dim fString As String
                         
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoCn.OpenSchema(adSchemaColumns)
       
    TableFlag = False
       
    Do Until frs.EOF = True
              
        If frs!Table_Name = TableName Then
            TableFlag = True
            Exit Do
        End If
      
        frs.MoveNext
   
    Loop

    frs.Close
    
    ' table does not exist
    If TableFlag = False Then Exit Sub

    fString = "DROP TABLE " & TableName
    adoCn.Execute fString

End Sub

Public Sub CustomerCreate()

    SQLString = "CREATE TABLE JCCustomer ( " & _
                        "[CustomerID] Counter, CONSTRAINT cusIDKey PRIMARY KEY ([CustomerID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "JCCustomer", "Name", "Char (50)", cn
    AddField "JCCustomer", "FullName", "Char (255)", cn
    AddField "JCCustomer", "CompanyName", "Char (255)", cn
    AddField "JCCustomer", "QBID", "Char (50)", cn
    AddField "JCCustomer", "FirstName", "Char (50)", cn
    AddField "JCCustomer", "LastName", "Char (50)", cn
    AddField "JCCustomer", "MidInit", "Char (50)", cn
    AddField "JCCustomer", "BillAddr1", "Char (50)", cn
    AddField "JCCustomer", "BillAddr2", "Char (50)", cn
    AddField "JCCustomer", "BillAddr3", "Char (50)", cn
    AddField "JCCustomer", "BillAddr4", "Char (50)", cn
    AddField "JCCustomer", "BillCity", "Char (50)", cn
    AddField "JCCustomer", "BillState", "Char (50)", cn
    AddField "JCCustomer", "BillZip", "Char (50)", cn
    AddField "JCCustomer", "ShipAddr1", "Char (50)", cn
    AddField "JCCustomer", "ShipAddr2", "Char (50)", cn
    AddField "JCCustomer", "ShipAddr3", "Char (50)", cn
    AddField "JCCustomer", "ShipAddr4", "Char (50)", cn
    AddField "JCCustomer", "ShipCity", "Char (50)", cn
    AddField "JCCustomer", "ShipState", "Char (50)", cn
    AddField "JCCustomer", "ShipZip", "Char (50)", cn
    AddField "JCCustomer", "QBTaxCode", "Char (50)", cn     ' TAX / NON
    AddField "JCCustomer", "QBTaxItem", "Char (50)", cn

End Sub

Public Sub JobCreate()

    SQLString = "CREATE TABLE JCJob ( " & _
                        "[JobID] Counter, CONSTRAINT jobIDKey PRIMARY KEY ([JobID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "JCJob", "Name", "Char (50)", cn
    AddField "JCJob", "FullName", "Char (255)", cn
    AddField "JCJob", "CompanyName", "Char (255)", cn
    AddField "JCJob", "QBID", "Char (50)", cn
    AddField "JCJob", "QBParentID", "Char(50)", cn
    AddField "JCJob", "ParentID", "Long", cn
    AddField "JCJob", "CityID", "Long", cn
    AddField "JCJob", "FirstName", "Char (50)", cn
    AddField "JCJob", "LastName", "Char (50)", cn
    AddField "JCJob", "MidInit", "Char (50)", cn
    AddField "JCJob", "BillAddr1", "Char (50)", cn
    AddField "JCJob", "BillAddr2", "Char (50)", cn
    AddField "JCJob", "BillAddr3", "Char (50)", cn
    AddField "JCJob", "BillAddr4", "Char (50)", cn
    AddField "JCJob", "BillCity", "Char (50)", cn
    AddField "JCJob", "BillState", "Char (50)", cn
    AddField "JCJob", "BillZip", "Char (50)", cn
    AddField "JCJob", "ShipAddr1", "Char (50)", cn
    AddField "JCJob", "ShipAddr2", "Char (50)", cn
    AddField "JCJob", "ShipAddr3", "Char (50)", cn
    AddField "JCJob", "ShipAddr4", "Char (50)", cn
    AddField "JCJob", "ShipCity", "Char (50)", cn
    AddField "JCJob", "ShipState", "Char (50)", cn
    AddField "JCJob", "ShipZip", "Char (50)", cn
    AddField "JCJob", "Status", "Char (50)", cn
    AddField "JCJob", "StartDate", "DateTime", cn
    AddField "JCJob", "EndDate", "DateTime", cn
    AddField "JCJob", "Description", "Char (50)", cn
    AddField "JCJob", "TypeName", "Char (50)", cn
    AddField "JCJob", "TypeListID", "Char (50)", cn
    AddField "JCJob", "JobStatus", "Byte", cn
    AddField "JCJob", "QBTaxCode", "Char (50)", cn
    AddField "JCJob", "Active", "Byte", cn
    AddField "JCJob", "Terms", "Char (50)", cn

End Sub

Public Sub QBAccountCreate()
    
    SQLString = "CREATE TABLE QBAccount ( " & _
                        "[QBAccountID] Counter, CONSTRAINT qbaIDKey PRIMARY KEY ([QBAccountID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "QBAccount", "Name", "Char (255)", cn
    AddField "QBAccount", "Description", "Char (255)", cn
    AddField "QBAccount", "AccountType", "Char (50)", cn
    AddField "QBAccount", "AccountNumber", "Long", cn
    AddField "QBAccount", "QBID", "Char (50)", cn

End Sub

Public Sub PRTimeSheetCreate()
    
    SQLString = "CREATE TABLE PRTimeSheet ( " & _
                        "[TimeSheetID] Counter, CONSTRAINT prtIDKey PRIMARY KEY ([TimeSheetID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "PRTimeSheet", "EmployeeID", "Long", cn
    AddField "PRTimeSheet", "JobID", "Long", cn
    AddField "PRTimeSheet", "CityID", "Long", cn
    AddField "PRTimeSheet", "DepartmentID", "Long", cn
    AddField "PRTimeSheet", "ItemID", "Long", cn
    AddField "PRTimeSheet", "Note", "Char (50)", cn
    AddField "PRTimeSheet", "SunHours", "Single", cn
    AddField "PRTimeSheet", "MonHours", "Single", cn
    AddField "PRTimeSheet", "TueHours", "Single", cn
    AddField "PRTimeSheet", "WedHours", "Single", cn
    AddField "PRTimeSheet", "ThuHours", "Single", cn
    AddField "PRTimeSheet", "FriHours", "Single", cn
    AddField "PRTimeSheet", "SatHours", "Single", cn
    AddField "PRTimeSheet", "TotalHours", "Single", cn
    AddField "PRTimeSheet", "HistID", "Long", cn
    AddField "PRTimeSheet", "BatchID", "Long", cn
    AddField "PRTimeSheet", "PEDate", "DateTime", cn
    AddField "PRTimeSheet", "CheckDate", "DateTime", cn
    AddField "PRTimeSheet", "WEDate", "DateTime", cn
    AddField "PRTimeSheet", "BillingRate", "Currency", cn
    AddField "PRTimeSheet", "QBInvID", "Char (50)", cn

    AddField "PRTimeSheet", "PWCraftID", "Long", cn
    AddField "PRTimeSheet", "PWUnionID", "Long", cn
    AddField "PRTimeSheet", "PWRegRate", "Currency", cn
    AddField "PRTimeSheet", "PWOvtRate", "Currency", cn
    AddField "PRTimeSheet", "PWFringeAmt", "Currency", cn

End Sub

Public Sub PRW2Create()

    SQLString = "CREATE TABLE PRW2 ( " & _
                        "[W2ID] Counter, CONSTRAINT prw2IDKey PRIMARY KEY ([W2ID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "PRW2", "TaxYear", "Long", cn
    AddField "PRW2", "EmployeeID", "Long", cn
    AddField "PRW2", "EmployeeNumber", "Long", cn
    AddField "PRW2", "BoxA_SSNumber", "Long", cn
    AddField "PRW2", "BoxB_FedID", "Char (15)", cn
    AddField "PRW2", "BoxC_ERName", "Char (50)", cn
    AddField "PRW2", "BoxC_ERAddr1", "Char (50)", cn
    AddField "PRW2", "BoxC_ERAddr2", "Char (50)", cn
    AddField "PRW2", "BoxC_ERCity", "Char (50)", cn
    AddField "PRW2", "BoxC_ERState", "Char (2)", cn
    AddField "PRW2", "BoxC_ERZip", "Char (10)", cn
    AddField "PRW2", "BoxD_ControlNumber", "Long", cn
    AddField "PRW2", "BoxE_EEFirstName", "Char (50)", cn
    AddField "PRW2", "BoxE_EELastName", "Char (50)", cn
    AddField "PRW2", "BoxE_EEMidInit", "Char (2)", cn
    AddField "PRW2", "BoxE_EEAddr1", "Char (50)", cn
    AddField "PRW2", "BoxE_EEAddr2", "Char (50)", cn
    AddField "PRW2", "BoxE_EECity", "Char (50)", cn
    AddField "PRW2", "BoxE_EEState", "Char (2)", cn
    AddField "PRW2", "BoxE_EEZip", "Char (10)", cn
    AddField "PRW2", "Box1_Wages", "Currency", cn
    AddField "PRW2", "Box2_FedTax", "Currency", cn
    AddField "PRW2", "Box3_SSWages", "Currency", cn
    AddField "PRW2", "Box4_SSTax", "Currency", cn
    AddField "PRW2", "Box5_MedWages", "Currency", cn
    AddField "PRW2", "Box6_MedTax", "Currency", cn
    AddField "PRW2", "Box7_SSTips", "Currency", cn
    AddField "PRW2", "Box8_AllocTips", "Currency", cn
    AddField "PRW2", "Box9_EIC", "Currency", cn
    AddField "PRW2", "Box10_DCBen", "Currency", cn
    AddField "PRW2", "Box11_NQPlans", "Currency", cn
    AddField "PRW2", "Box12A_ID", "Long", cn
    AddField "PRW2", "Box12A_Code", "Char (1)", cn
    AddField "PRW2", "Box12A_Amount", "Currency", cn
    AddField "PRW2", "Box12B_ID", "Long", cn
    AddField "PRW2", "Box12B_Code", "Char (1)", cn
    AddField "PRW2", "Box12B_Amount", "Currency", cn
    AddField "PRW2", "Box12C_ID", "Long", cn
    AddField "PRW2", "Box12C_Code", "Char (1)", cn
    AddField "PRW2", "Box12C_Amount", "Currency", cn
    AddField "PRW2", "Box12D_ID", "Long", cn
    AddField "PRW2", "Box12D_Code", "Char (1)", cn
    AddField "PRW2", "Box12D_Amount", "Currency", cn
    AddField "PRW2", "Box13_StatEmp", "Byte", cn
    AddField "PRW2", "Box13_RetirePlan", "Byte", cn
    AddField "PRW2", "Box13_3rdParty", "Byte", cn
    AddField "PRW2", "Box14A_ID", "Long", cn
    AddField "PRW2", "Box14A_Desc", "Char (10)", cn
    AddField "PRW2", "Box14A_Amount", "Currency", cn
    AddField "PRW2", "Box14B_ID", "Long", cn
    AddField "PRW2", "Box14B_Desc", "Char (10)", cn
    AddField "PRW2", "Box14B_Amount", "Currency", cn
    AddField "PRW2", "Box14C_ID", "Long", cn
    AddField "PRW2", "Box14C_Desc", "Char (10)", cn
    AddField "PRW2", "Box14C_Amount", "Currency", cn
    AddField "PRW2", "Box14D_ID", "Long", cn
    AddField "PRW2", "Box14D_Desc", "Char (10)", cn
    AddField "PRW2", "Box14D_Amount", "Currency", cn
    AddField "PRW2", "Box15A_State", "Char (2)", cn
    AddField "PRW2", "Box15A_StateID", "Char (20)", cn
    AddField "PRW2", "Box16A_StateWages", "Currency", cn
    AddField "PRW2", "Box17A_StateTax", "Currency", cn
    AddField "PRW2", "Box15B_State", "Char (2)", cn
    AddField "PRW2", "Box15B_StateID", "Char (20)", cn
    AddField "PRW2", "Box16B_StateWages", "Currency", cn
    AddField "PRW2", "Box17B_StateTax", "Currency", cn
    AddField "PRW2", "Void", "Byte", cn
    AddField "PRW2", "Skip", "Byte", cn

End Sub

Public Sub PRW2CityCreate()

    SQLString = "CREATE TABLE PRW2City ( " & _
                        "[W2CityID] Counter, CONSTRAINT prw2cIDKey PRIMARY KEY ([W2CityID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "PRW2City", "W2ID", "Long", cn
    AddField "PRW2City", "TaxYear", "Long", cn
    AddField "PRW2City", "CityID", "Long", cn
    AddField "PRW2City", "CityName", "Char (20)", cn
    AddField "PRW2City", "CityWage", "Currency", cn
    AddField "PRW2City", "CityTax", "Currency", cn
    AddField "PRW2City", "StateID", "Long", cn
    AddField "PRW2City", "SDTax", "Byte", cn
    AddField "PRW2City", "Courtesy", "Byte", cn

End Sub

Public Sub PRW2StateCreate()
    
    SQLString = "CREATE TABLE PRW2State ( " & _
                        "[W2StateID] Counter, CONSTRAINT prw2sIDKey PRIMARY KEY ([W2StateID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "PRW2State", "W2ID", "Long", cn
    AddField "PRW2State", "TaxYear", "Long", cn
    AddField "PRW2State", "StateID", "Long", cn
    AddField "PRW2State", "ERStateID", "Char (20)", cn
    AddField "PRW2State", "StateWage", "Currency", cn
    AddField "PRW2State", "StateTax", "Currency", cn

End Sub

Public Sub EmpRelatedCreate()

    SQLString = "CREATE TABLE PREmpRelated ( " & _
                        "[EmpRelatedID]       Counter, CONSTRAINT emprelIDKey PRIMARY KEY ([EmpRelatedID]), " & _
                        "[EmployeeID]   Long, CONSTRAINT emprelNumberKey UNIQUE ([EmployeeID]) ) "
    
    cn.Execute SQLString

    AddField "PREmpRelated", "Comment", "Memo", cn

End Sub

Public Sub NotesCreate()

    SQLString = "CREATE TABLE Notes ( " & _
                        "[NoteID] Counter, CONSTRAINT nteIDKey PRIMARY KEY ([NoteID]) ) "
                        
    cn.Execute SQLString

    AddField "Notes", "NoteType", "Byte", cn
    AddField "Notes", "NoteCat", "Byte", cn
    AddField "Notes", "RelatedID", "Long", cn
    AddField "Notes", "Subject", "Char (15)", cn
    AddField "Notes", "User", "Char (8)", cn
    AddField "Notes", "DateTm", "DateTime", cn
    AddField "Notes", "Notation", "Memo", cn

End Sub

Public Sub QBUpdateCreate()
    
    SQLString = "CREATE TABLE QBUpdate ( " & _
                        "[QBUpdateID] Counter, CONSTRAINT qbupIDKey PRIMARY KEY ([QBUpdateID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "QBUpdate", "Category", "Byte", cn
    AddField "QBUpdate", "Post", "Char (1)", cn
    AddField "QBUpdate", "PerJob", "Byte", cn
    AddField "QBUpdate", "Title", "Char (30)", cn
    AddField "QBUpdate", "Type", "Byte", cn
    AddField "QBUpdate", "RelatedID", "Long", cn
    AddField "QBUpdate", "QBID", "Char (50)", cn
    AddField "QBUpdate", "DebitAmount", "Currency", cn
    AddField "QBUpdate", "CreditAmount", "Currency", cn

End Sub

Public Sub PRCountyCreate()

    SQLString = "CREATE TABLE PRCounty ( " & _
                        "[CountyID] Counter, CONSTRAINT ctyIDKey PRIMARY KEY ([CountyID]) ) "
                        
    cnDes.Execute SQLString
                        
    AddField "PRCounty", "CountyName", "Char (50)", cnDes
    AddField "PRCounty", "ShortName", "Char (10)", cnDes
    AddField "PRCounty", "StateID", "Long", cnDes
    AddField "PRCounty", "SalesTaxRate", "Double", cnDes

End Sub

Public Sub GLFFSchedCreate()

    SQLString = "CREATE TABLE GLFFSched ( " & _
                        "[FFSchedID] Counter, CONSTRAINT glffsIDKey PRIMARY KEY ([FFSchedID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "GLFFSched", "GlobalID", "Long", cn
    AddField "GLFFSched", "Account", "Long", cn
    AddField "GLFFSched", "SortOrder", "Long", cn
    AddField "GLFFSched", "PercentBase", "Long", cn
    AddField "GLFFSched", "PrintTab", "Byte", cn
    AddField "GLFFSched", "LineFeeds", "Byte", cn
    AddField "GLFFSched", "SignReverse", "Byte", cn
    AddField "GLFFSched", "AltDesc", "Char (50)", cn
    AddField "GLFFSched", "ReportID", "Byte", cn

End Sub

Public Sub PRW4Create()

    SQLString = "CREATE TABLE PRW4 ( " & _
                        "[W4ID] Counter, CONSTRAINT w4IDKey PRIMARY KEY ([W4ID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "PRW4", "EmployeeID", "Long", cn
    AddField "PRW4", "TwoJobs", "Byte", cn
    AddField "PRW4", "FilingType", "Byte", cn
    AddField "PRW4", "Dependents", "Byte", cn
    AddField "PRW4", "DependentsOther", "Byte", cn
    AddField "PRW4", "OtherIncome", "Currency", cn
    AddField "PRW4", "Deductions", "Currency", cn
    AddField "PRW4", "ExtraWH", "Currency", cn

End Sub


