VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquate"
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

Public Stmt As Byte
Public Sched As Byte
Public NonComp As Byte
Public Comp As Byte
Public Regular As Byte
Public Branch As Byte
Public Consol As Byte
Public Budget As Byte
Public PrtBSOnly As Byte
Public PrtISOnly As Byte
Public PrtBoth As Byte
Public PgTwips As Long
Public RecAdd As Boolean
Public RecPut As Boolean
Public FormAdd As Byte
Public FormEdit As Byte
Public FormDel As Byte
Public Portrait As Byte
Public Landscape As Byte

Public GlobalTypeNudge As Byte

Public fmtAmount As Byte
Public fmtString As Byte



