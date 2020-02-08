VERSION 5.00
Begin VB.Form frmDescriptions 
   Caption         =   " General Ledger Descriptions"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MsgBox "load"
End Sub

Private Sub datamover()
    On Error GoTo glerr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("glDescriptions")
    Dim db2 As Database
    Dim rs2 As Recordset
    Set db2 = OpenDatabase("\balint\data\glDesc.mdb")
    Set rs2 = db2.OpenRecordset("GLDescription")
    Do Until rs2.EOF = True
        rs.AddNew
        rs.Fields("number") = rs2.Fields("number")
        rs.Fields("Description") = rs2.Fields("description")
        rs.Update
        rs2.MoveNext
    Loop
    MsgBox rs.RecordCount
    MsgBox rs2.RecordCount
    Exit Sub
glerr:
    MsgBox Error(Err.Number)
End Sub
