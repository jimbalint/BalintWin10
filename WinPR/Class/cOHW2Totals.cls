VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOHW2Totals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RWTotalCount As Integer
Public RWCount As Integer
Public ROCount As Integer

Public Box1_Wages As Currency
Public Box2_FedTax As Currency
Public Box3_SSWages As Currency
Public Box4_SSTax As Currency
Public Box5_MedWages As Currency
Public Box6_MedTax As Currency
Public Box7_SSTips As Currency
Public Box8_AllocTips As Currency
Public Box9_EIC As Currency
Public Box10_DCBen As Currency
Public Box11_NQPlans As Currency
Public CodeA As Currency
Public CodeB As Currency
Public CodeD As Currency
Public CodeE As Currency
Public CodeF As Currency
Public CodeG As Currency
Public CodeH As Currency
Public CodeM As Currency
Public CodeN As Currency
Public CodeR As Currency
Public CodeS As Currency
Public CodeT As Currency
Public CodeW As Currency
Public CodeQ As Currency
Public CodeZ As Currency
Public CodeDD As Currency
Public CodeC As Currency
Public IncTax3rdSick As Currency
Public CodeV As Currency
Public CodeY As Currency
Public CodeAA As Currency
Public CodeBB As Currency
Public CodeEE As Currency
Public CodeFF As Currency
Public CodeGG As Currency
Public CodeHH As Currency
Public RetireAmt As Currency

Public Sub Clear()
    RWCount = 0
    ROCount = 0
    Box1_Wages = 0
    Box2_FedTax = 0
    Box3_SSWages = 0
    Box4_SSTax = 0
    Box5_MedWages = 0
    Box6_MedTax = 0
    Box7_SSTips = 0
    Box8_AllocTips = 0
    Box9_EIC = 0
    Box10_DCBen = 0
    Box11_NQPlans = 0
    CodeA = 0
    CodeB = 0
    CodeD = 0
    CodeE = 0
    CodeF = 0
    CodeG = 0
    CodeH = 0
    CodeM = 0
    CodeN = 0
    CodeR = 0
    CodeS = 0
    CodeT = 0
    CodeW = 0
    CodeQ = 0
    CodeZ = 0
    CodeDD = 0
    CodeC = 0
    IncTax3rdSick = 0
    CodeV = 0
    CodeY = 0
    CodeAA = 0
    CodeBB = 0
    CodeEE = 0
    CodeFF = 0
    CodeGG = 0
    CodeHH = 0
    RetireAmt = 0
End Sub


