VERSION 5.00
Begin VB.Form frmUpdBatch 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmUpdBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xDB As New XArrayDB
         
Private Sub Form_Load()
         
         ' get the batch number if necessary
         If Period = 0 Then
            GLBatch.OpenRS
         
            If Not GLBatch.GetBatch(BatchNum) Then
               MsgBox "Update failed! Batch #: " & BatchNum & " does not exist!", vbCritical + vbOKOnly, "Update Amounts"
               End
            End If
         End If
         
         xDB.ReDim 0, 5, 0, 0
         xDB(1, 0) = GLCompany.Name & " Clear and Update"
         xDB(2, 0) = "Fiscal Year: " & GLBatch.FiscalYear & " " & _
                     "Start Period: " & GLBatch.Period & " " & _
                     "End Period: " & GLBatch.Period
         xDB(3, 0) = " "
         xDB(4, 0) = String(40, "=")
         xDB(5, 0) = " "

         booX = False  ' don't delete history - clear and reupdate it

         Set uDB = ClearGLAmount(GLBatch.FiscalYear, _
                                 GLBatch.FiscalYear, _
                                 GLBatch.Period, _
                                 GLBatch.Period, _
                                 booX)

         xDBAssign

         Set uDB = ClearGLBudget(GLBatch.FiscalYear, _
                                 GLBatch.FiscalYear, _
                                 GLBatch.Period, _
                                 GLBatch.Period)

         xDBAssign

         ' create suspense account
         Set uDB = UpdateGLAmount(GLBatch.FiscalYear, _
                                  GLBatch.FiscalYear, _
                                  GLBatch.Period, _
                                  GLBatch.Period, _
                                  0, _
                                  CompanyID)

         xDBAssign


         Set uDB = MathUpdate(GLBatch.FiscalYear, _
                              GLBatch.FiscalYear, _
                              GLBatch.Period, _
                              GLBatch.Period)

         xDBAssign

         Set frmResults = New frmResults
         frmResults.lblCompanyName = GLCompany.Name
         frmResults.lblMsg1 = "Clear and Update GL Amounts"
         frmResults.lblMsg2 = ""
         frmResults.lblMsg3 = ""
         For i = 1 To xDB.UpperBound(1)
             frmResults.List1.AddItem xDB(i, 0)
         Next i
         frmResults.Show vbModal

         GoBack

End Sub
Private Sub xDBAssign()

Dim i As Integer
Dim j As Integer
    
    For i = 1 To uDB.UpperBound(1)
        xDB.AppendRows
        j = xDB.UpperBound(1)
        xDB(j, 0) = uDB(i, 0)
    Next i

End Sub

