VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvEquate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GlobalTypeTruck As Byte
Public GlobalTypeTrailer As Byte
Public GlobalTypeDriver As Byte
Public GlobalTypeTerms As Byte
Public GlobalTypeComment As Byte
Public GlobalTypeInvNumber As Byte
Public GlobalTypeQBSetup As Byte
Public GlobalTypeInvPrinter As Byte
Public GlobalTypeInvMessage As Byte
Public GlobalTypeSalesTax As Byte       ' does this QB company have sales tax?
Public GlobalTypeVAdj As Byte

Public IBMCPI10 As String
Public IBMCPI12 As String
Public IBMCPI17 As String
Public IBMDblWide As String
