VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEquate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public OpAdd As Byte
Public OpSub As Byte
Public OpPct As Byte
Public OpMult As Byte
Public OpDiv As Byte

Public NoteTypeEE As Byte
Public NoteTypeER As Byte

Public ColAvg As Byte
Public ColProj As Byte
Public ColMultiply As Byte
Public ColAdd As Byte
Public ColSubtract As Byte
Public ColDivide As Byte
Public ColCurrPd As Byte
Public ColPriorPd As Byte
Public ColAllPd As Byte
Public ColYTD As Byte
Public ColCustom As Byte

Private mStmt As Byte
Private mSched As Byte
Private mNonComp As Byte
Private mComp As Byte
Private mRegular As Byte
Private mBranch As Byte
Private mConsol As Byte
Private mBudget As Byte
Private mPrtBSOnly As Byte
Private mPrtISOnly As Byte
Private mPrtBoth As Byte
Private mPgTwips As Long
Private mRecAdd As Boolean
Private mRecPut As Boolean
Private mFormAdd As Byte
Private mFormEdit As Byte
Private mFormDel As Byte
Private mPortrait As Byte
Private mLandscape As Byte


Public Property Let Stmt(ByVal bytNew As Byte)
    mStmt = bytNew
End Property

Public Property Get Stmt() As Byte
    Stmt = mStmt
End Property


Public Property Let Sched(ByVal bytNew As Byte)
    mSched = bytNew
End Property

Public Property Get Sched() As Byte
    Sched = mSched
End Property



Public Property Let NonComp(ByVal bytNew As Byte)
    mNonComp = bytNew
End Property

Public Property Get NonComp() As Byte
    NonComp = mNonComp
End Property


Public Property Let Comp(ByVal bytNew As Byte)
    mComp = bytNew
End Property

Public Property Get Comp() As Byte
    Comp = mComp
End Property


Public Property Let Regular(ByVal bytNew As Byte)
    mRegular = bytNew
End Property

Public Property Get Regular() As Byte
    Regular = mRegular
End Property


Public Property Let Branch(ByVal bytNew As Byte)
    mBranch = bytNew
End Property

Public Property Get Branch() As Byte
    Branch = mBranch
End Property


Public Property Let Consol(ByVal bytNew As Byte)
    mConsol = bytNew
End Property

Public Property Get Consol() As Byte
    Consol = mConsol
End Property


Public Property Let Budget(ByVal bytNew As Byte)
    mBudget = bytNew
End Property

Public Property Get Budget() As Byte
    Budget = mBudget
End Property


Public Property Let PrtBSOnly(ByVal bytNew As Byte)
    mPrtBSOnly = bytNew
End Property

Public Property Get PrtBSOnly() As Byte
    PrtBSOnly = mPrtBSOnly
End Property


Public Property Let PrtISOnly(ByVal bytNew As Byte)
    mPrtISOnly = bytNew
End Property

Public Property Get PrtISOnly() As Byte
    PrtISOnly = mPrtISOnly
End Property


Public Property Let PrtBoth(ByVal bytNew As Byte)
    mPrtBoth = bytNew
End Property

Public Property Get PrtBoth() As Byte
    PrtBoth = mPrtBoth
End Property


Public Property Let PgTwips(ByVal lngNew As Long)
    mPgTwips = lngNew
End Property

Public Property Get PgTwips() As Long
    PgTwips = mPgTwips
End Property


Public Property Let RecAdd(ByVal booNew As Boolean)
    mRecAdd = booNew
End Property

Public Property Get RecAdd() As Boolean
    RecAdd = mRecAdd
End Property


Public Property Let RecPut(ByVal booNew As Boolean)
    mRecPut = booNew
End Property

Public Property Get RecPut() As Boolean
    RecPut = mRecPut
End Property


Public Property Let FormAdd(ByVal bytNew As Byte)
    mFormAdd = bytNew
End Property

Public Property Get FormAdd() As Byte
    FormAdd = mFormAdd
End Property


Public Property Let FormEdit(ByVal bytNew As Byte)
    mFormEdit = bytNew
End Property

Public Property Get FormEdit() As Byte
    FormEdit = mFormEdit
End Property


Public Property Let FormDel(ByVal bytNew As Byte)
    mFormDel = bytNew
End Property

Public Property Get FormDel() As Byte
    FormDel = mFormDel
End Property

Public Property Let Portrait(ByVal bytNew As Byte)
    mPortrait = bytNew
End Property

Public Property Get Portrait() As Byte
    Portrait = mPortrait
End Property

Public Property Let LandScape(ByVal bytNew As Byte)
    mLandscape = bytNew
End Property

Public Property Get LandScape() As Byte
    LandScape = mLandscape
End Property

