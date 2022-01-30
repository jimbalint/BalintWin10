VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPREquate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ItemTypeRegPay As Byte
Public ItemTypeOvtPay As Byte
Public ItemTypeOE As Byte
Public ItemTypeDED As Byte
Public ItemTypeSDTax As Byte
Public ItemTypeOtherTax As Byte
Public ItemTypeWorkComp As Byte
Public ItemTypeShiftDiff As Byte
Public ItemType1099Pay As Byte
Public ItemTypeDirDepDed As Byte
Public ItemTypeSick As Byte
Public ItemTypeVacation As Byte
Public ItemTypeNotInNet As Byte
Public ItemTypeRegTax As Byte
Public ItemType3PSickPay As Byte
Public ItemTypeEIC As Byte

Public qbItem_RegPay As Byte
Public qbItem_OvtPay As Byte
Public qbItem_OE As Byte
Public qbItem_SSTax As Byte
Public qbItem_MEDTax As Byte
Public qbItem_FWTTax As Byte
Public qbItem_SWTTax As Byte
Public qbItem_CWTTax As Byte
Public qbItem_DED As Byte
Public qbItem_Net As Byte
Public qbItem_ERSS As Byte
Public qbItem_ERMED As Byte
Public qbItem_ERFUN As Byte
Public qbItem_ERSUN As Byte
Public qbItem_ERWkc As Byte
Public qbItem_AccrueSS As Byte
Public qbItem_AccrueMED As Byte
Public qbItem_AccrueFUN As Byte
Public qbItem_AccrueSUN As Byte
Public qbItem_AccrueWkc As Byte

Public qbJobStatus_Awarded As Byte
Public qbJobStatus_Closed As Byte
Public qbJobStatus_InProgress As Byte
Public qbJobStatus_None As Byte
Public qbJobStatus_NotAwarded As Byte
Public qbJobStatus_Pending As Byte

Public DistTypeReg As Byte
Public DistTypeOT As Byte
Public DistTypeItem As Byte

Public BasisAmount As Byte
Public BasisPercent As Byte
Public BasisHourly As Byte
Public BasisNet As Byte
Public BasisExemptions As Byte

Public DirDepTypeChecking As Byte
Public DirDepTypeSavings As Byte

Public DirDepBasisAmt As Byte
Public DirDepBasisPct As Byte
Public DirDepBasisNet As Byte

Public RecTypeAdjust As Byte
Public RecTypeDepartment As Byte
Public RecTypeEERate As Byte
Public RecTypeEmployee As Byte
Public RecTypeEmployer As Byte
Public RecTypeHistory As Byte
Public RecTypeItem As Byte
Public RecTypeList As Byte
Public RecTypePhone As Byte

Public WageGross As Byte
Public WageFed As Byte
Public WageSS As Byte
Public WageMed As Byte
Public WageState As Byte
Public WageCity As Byte
Public WageSUN As Byte
Public WageFUN As Byte

Public TaxFed As Byte
Public TaxSS As Byte
Public TaxMed As Byte
Public TaxState As Byte
Public TaxCity As Byte

Public TipsSS As Byte
Public TipsMed As Byte

Public EEListsPhone As Byte
Public EEListsDptRate As Byte
Public EEListsPRNotes As Byte

Public GlobalTypeRaceCode As Byte
Public GlobalTypeEducationLevel As Byte
Public GlobalTypeContact As Byte
Public GlobalTypeShiftCode As Byte
Public GlobalTypeTerminationCode As Byte

Public GlobalTypeSSMax As Byte
Public GlobalTypeSSPct As Byte
Public GlobalTypeMEDPct As Byte
Public GlobalTypeMEDAddPct As Byte
Public GlobalTypeMEDAddAmt As Byte
Public GlobalTypeFWTAllow As Byte
Public GLobalTypeOHAllow As Byte
Public GlobalTypeFUNMax As Byte
Public GlobalTypeFUNPct As Byte
Public GlobalTypeSUNMax As Byte
Public GlobalTypeWkcCat As Byte
Public GlobalTypeEICMaxAdv As Byte
Public GlobalTypeEICMaxWage As Byte

Public GlobalTypeW2Box12 As Byte
Public GlobalTypeW2Box14 As Byte
Public GlobalTypePaysPerYear As Byte

Public GlobalTypeNudge As Byte

Public GlobalType941Part4 As Byte
Public GlobalType941Part5 As Byte
Public GlobalType941PaidPrep As Byte

Public GlobalTypeBUCForm As Byte
Public GlobalTypePRCheckPrefix As Byte
Public GlobalTypeOHSDTaxAllow As Byte

Public GlobalTypeW3A As Byte
Public GlobalTypeW3B As Byte
Public GlobalTypeW3C As Byte
Public GlobalTypeW3D As Byte
Public GlobalTypeW3E As Byte

Public GlobalTypeQBPRChk As Byte
Public GlobalTypeDeductBasis As Byte
Public GlobalTypeDirDepFolder As Byte

Public GlobalTypeQB_EE_Wage As Byte
Public GlobalTypeQB_EE_Tax As Byte
Public GlobalTypeQB_EE_Ded As Byte
Public GlobalTypeQB_ER_Tax As Byte


Public GlobalTypeQB_File As Byte ' *** not used ***

Public GlobalTypeQB_Register As Byte

Public GlobalTypeScreenDefault As Byte

Public GlobalTypeQBInv As Byte

' link TimeSheet Week Ended to PRBatch
Public GlobalTypePRBatchWE As Byte

Public GlobalTypePWCraft As Byte
Public GlobalTypePWCounty As Byte
Public GlobalTypePWUnion As Byte
Public GlobalTypePWWage As Byte

Public GlobalTypeQBPayFED As Byte
Public GlobalTypeQBPayState As Byte
Public GlobalTypeQBPayCity As Byte
Public GlobalTypeQBPayWKC As Byte
Public GlobalTypeQBPayFUN As Byte
Public GlobalTypeQBPaySUN As Byte
Public GlobalTypeQBPayItem As Byte
Public GlobalTypeQBPayDeductItem As Byte
Public GlobalTypeQBPayNetPay As Byte
Public GlobalTypeQBPayGrossPay As Byte
Public GlobalTypeQBPayCompany As Byte

Public GlobalTypeCompanyOption As Byte

Public GlobalTypeGLFFSched As Byte
Public GlobalTypeGLFFColumn As Byte
Public GlobalTypeGLFFSetup As Byte

Public GlobalTypeUnionEmployee As Byte
Public GlobalTypeUnionJob As Byte

Public GlobalTypeOtherStateID As Byte

' 2019-12-14 OH SWT multiplier
Public GlobalTypeOHMultiplier As Double

' 2020-12-12
Public GlobalTypeLastBackUp As Double

Public RangeTypeBatch As Byte
Public RangeTypeMonths As Byte
Public RangeTypePEDate As Byte

Public EICTypeSingle As Byte
Public EICTypeMarriedWith As Byte
Public EICTypeMarriedWO As Byte

Public x1099Reg As Byte
Public x1099Inc As Byte

Public GLTypeEmployee As Byte
Public GLTypeDept As Byte
Public GLTypeCompany As Byte

Public GLItemTypeOE As Byte
Public GLItemTypeDed As Byte
Public GLItemTypeSSTax As Byte
Public GLItemTypeSSMatch As Byte
Public GLItemTypeMedTax As Byte
Public GLItemTypeMedMatch As Byte
Public GLItemTypeFWTTax As Byte
Public GLItemTypeSWTTax As Byte
Public GLItemTypeCWTTax As Byte
Public GLItemTypeFUN As Byte
Public GLItemTypeSUN As Byte
Public GLItemTypeGross As Byte
Public GLItemTypeNet As Byte
Public GLItemTypeSSExp As Byte
Public GLItemTypeMEDExp As Byte
Public GLItemTypeFUNExp As Byte
Public GLItemTypeSUNExp As Byte
Public GLItemTypeWkcExp As Byte
Public GLItemTypeRegPay As Byte
Public GLItemTypeOvtPay As Byte
Public GLItemTypeWkcTax As Byte

Public SortOrderNumber As Byte
Public SortOrderName As Byte
Public SortOrderDeptNumber As Byte
Public SortOrderDeptName As Byte

Public ActionAdd As Byte
Public ActionEdit As Byte
Public ActionDelete As Byte

Public CheckTypeBlankStock As Byte
Public CheckTypePrePrintedA As Byte
Public CheckTypePrePrintedB As Byte
Public CheckTypePrePrintedC As Byte

Public QBError_NoFileOpen As Double
Public QBError_QBBeginSession As Double

Public PeriodTypeMonth As Byte
Public PeriodTypeQuarter As Byte
Public PeriodTypeYear As Byte
Public PeriodTypePay As Byte

Public PayTypeCheck As Byte
Public PayTypeEFT As Byte

Public GlobalTypeOHW2Company As Byte
Public GlobalTypeOHW2Contact As Byte
Public GlobalTypeOHW2Submit As Byte




