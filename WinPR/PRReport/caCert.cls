VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "caCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RegHrs As Byte
Public OvtHrs As Byte
Public RegRate As Byte
Public OvtRate As Byte
Public RegGross As Byte
Public OvtGross As Byte
Public TotGross As Byte
Public RegTaxGross As Byte
Public OvtTaxGross As Byte
Public TotTaxGross As Byte
Public Net As Byte
Public SunRegHrs As Byte
Public MonRegHrs As Byte
Public TueRegHrs As Byte
Public WedRegHrs As Byte
Public ThuRegHrs As Byte
Public FriRegHrs As Byte
Public SatRegHrs As Byte
Public SunOvtHrs As Byte
Public MonOvtHrs As Byte
Public TueOvtHrs As Byte
Public WedOvtHrs As Byte
Public ThuOvtHrs As Byte
Public FriOvtHrs As Byte
Public SatOvtHrs As Byte
Public SSTax As Byte
Public MedTax As Byte
Public FWTTax As Byte
Public SWTTax As Byte
Public CWTTax As Byte
Public TotalTax As Byte

Public Sub AssignValues()

    RegHrs = 1
    OvtHrs = 2
    RegRate = 3
    OvtRate = 4
    RegGross = 5
    OvtGross = 6
    TotGross = 7
    RegTaxGross = 8
    OvtTaxGross = 9
    TotTaxGross = 10
    Net = 11
    SunRegHrs = 12
    MonRegHrs = 13
    TueRegHrs = 14
    WedRegHrs = 15
    ThuRegHrs = 16
    FriRegHrs = 17
    SatRegHrs = 18
    SunOvtHrs = 19
    MonOvtHrs = 20
    TueOvtHrs = 21
    WedOvtHrs = 22
    ThuOvtHrs = 23
    FriOvtHrs = 24
    SatOvtHrs = 25
    SSTax = 26
    MedTax = 27
    FWTTax = 28
    SWTTax = 29
    CWTTax = 30
    TotalTax = 31

End Sub
