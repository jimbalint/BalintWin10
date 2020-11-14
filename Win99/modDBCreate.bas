Attribute VB_Name = "modDBCreate"
Option Explicit

Dim FieldOrderNum As Integer

Public Sub Form99Create()
    
    SQLString = "CREATE TABLE Form99 ( " & _
                        "[FormID] Counter, CONSTRAINT frmIDKey PRIMARY KEY ([FormID]) ) "
                        
    cn99.Execute SQLString
    
    AddField "Form99", "FormType", "char (10)", cn99
    AddField "Form99", "TaxYear", "Long", cn99
    AddField "Form99", "FormsPerPg", "Byte", cn99
    AddField "Form99", "VersionNum", "char (10)", cn99
    AddField "Form99", "FormVert1", "Long", cn99
    AddField "Form99", "FormVert2", "Long", cn99
    AddField "Form99", "FormVert3", "Long", cn99
    AddField "Form99", "FormVert4", "Long", cn99
    
End Sub

Public Sub Field99Create()
    
    SQLString = "CREATE TABLE Field99 ( " & _
                        "[FieldID] Counter, CONSTRAINT fldIDKey PRIMARY KEY ([FieldID]) ) "
                        
    cn99.Execute SQLString
    
    AddField "Field99", "TaxYear", "Long", cn99
    AddField "Field99", "FormType", "char (10)", cn99
    AddField "Field99", "BoxName", "char (50)", cn99
    AddField "Field99", "FieldOrder", "Long", cn99
    AddField "Field99", "FieldTitle", "char (50)", cn99
    AddField "Field99", "FieldFormat", "Byte", cn99
    AddField "Field99", "HorzPosn", "Long", cn99
    AddField "Field99", "VertPosn", "Long", cn99
    AddField "Field99", "QuickEntry", "Byte", cn99
    
End Sub

Public Sub Payee99Create()
    
    SQLString = "CREATE TABLE Payee99 ( " & _
                        "[PayeeID] Counter, CONSTRAINT peeIDKey PRIMARY KEY ([PayeeID]) ) "
                        
    cn.Execute SQLString
    
    AddField "Payee99", "PayeeName", "char (50)", cn
    AddField "Payee99", "PayeeNumber", "Long", cn
    AddField "Payee99", "Address", "char (50)", cn
    AddField "Payee99", "CSZ", "char (50)", cn
    AddField "Payee99", "FederalID", "char (15)", cn
    AddField "Payee99", "AccountNumber", "char (50)", cn
    AddField "Payee99", "Comment", "char (50)", cn
    AddField "Payee99", "Inactive", "Byte", cn
    
End Sub

Public Sub Detail99Create()
    
    SQLString = "CREATE TABLE Detail99 ( " & _
                        "[DetailID] Counter, CONSTRAINT dtlIDKey PRIMARY KEY ([DetailID]) ) "
                        
    cn.Execute SQLString
    
    AddField "Detail99", "PayeeID", "Long", cn
    AddField "Detail99", "TaxYear", "Long", cn
    AddField "Detail99", "FormType", "char (10)", cn
    AddField "Detail99", "BoxName", "char (50)", cn
    AddField "Detail99", "FieldValue", "char (20)", cn
    
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
Dim FString As String
                         
                         
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

    FString = "DROP TABLE " & TableName
    adoCn.Execute FString

End Sub

Public Sub ClearFormRecs(ByVal FormType As String, ByVal TaxYear As Long)

    SQLString = "DELETE * FROM Form99 WHERE TaxYear = " & TaxYear & _
                " AND FormType = '" & FormType & "'"
    cn99.Execute SQLString
    
    SQLString = "DELETE * FROM Field99 WHERE TaxYear = " & TaxYear & _
                " AND FormType = '" & FormType & "'"
    cn99.Execute SQLString

End Sub
Public Sub Create2020Forms(ByVal jFormType As String)
    
Dim TaxYear As Long
Dim HorzPosn1, HorzPosn2, HorzPosn3, Tab1, Tab2 As Integer
Dim VertSpacing, VertPosn As Integer
Dim FormID As Long
Dim FormType As String

    TaxYear = 2020

    Form99.OpenRS
    Form99.Clear
    
    
    ' ================================================================================================
    ' MISC
    
    If jFormType = "MISC" Then
    
        ClearFormRecs "MISC", TaxYear
        
        Form99.FormType = "MISC"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 1000
        Form99.FormVert2 = 8930
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "MISC"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Rents", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "2", "Royalties", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 220
        Field99Add TaxYear, FormType, "3", "Other Income", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 50
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "5", "Fishing boat proceeds", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "6", "Medical and health care payments", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "7", "Payer made direct sales", Equate.fmtString, HorzPosn1 + 1430, VertPosn, 7
        Field99Add TaxYear, FormType, "8", "Sub pmts in lieu of div or int", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "9", "Corp insurance proceeds", Equate.fmtAmount, HorzPosn1, VertPosn, 9
        Field99Add TaxYear, FormType, "10", "Gross proceeds to attny", Equate.fmtAmount, HorzPosn2, VertPosn, 10
        
        VertPosn = VertPosn + VertSpacing
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "12", "Section 409A deferrals", Equate.fmtAmount, HorzPosn2, VertPosn, 11
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "13", "Excess golden parachute", Equate.fmtAmount, HorzPosn1, VertPosn, 12
        Field99Add TaxYear, FormType, "14", "Nonqualified deferred compensation", Equate.fmtAmount, HorzPosn2, VertPosn, 13
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "15a", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 14
        Field99Add TaxYear, FormType, "16a", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 15
        Field99Add TaxYear, FormType, "17a", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 16
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "15b", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 17
        Field99Add TaxYear, FormType, "16b", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 18
        Field99Add TaxYear, FormType, "17b", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 19
    
    End If
    
    ' ================================================================================================
    ' NEC
    
    If jFormType = "NEC" Then
    
        ClearFormRecs "NEC", TaxYear
        
        Form99.FormType = "NEC"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 1000
        Form99.FormVert2 = 8930
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "NEC"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 220
        Field99Add TaxYear, FormType, "1", "Nonemployee compensation", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 50
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn1 + 1430, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 3
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "5a", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        Field99Add TaxYear, FormType, "6a", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 4
        Field99Add TaxYear, FormType, "7a", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 5
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "5b", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 6
        Field99Add TaxYear, FormType, "6b", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 7
        Field99Add TaxYear, FormType, "7b", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 8
    
    End If
    
    
    ' ================================================================================================
    ' 1096
    
    If jFormType = "1096" Then
    
         ClearFormRecs "1096", TaxYear
         
         Form99.FormType = "1096"
         Form99.TaxYear = TaxYear
         Form99.FormsPerPg = 1
         Form99.FormVert1 = 1000
         Form99.FormVert2 = 8950
         Form99.Save (Equate.RecAdd)
         FormID = Form99.FormID
         
         ' TaxYear, FormType,  BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
         Tab1 = 920
         Tab2 = 500
         FieldOrderNum = 0
         FormType = "1096"
         
         VertPosn = 850
         Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1450
         Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1700
         Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2150
         Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2595
         Field99Add TaxYear, FormType, "Contact", "Contact", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Phone", "Phone", Equate.fmtString, Tab2 + 4750, VertPosn, 0
             
         VertPosn = 3110
         Field99Add TaxYear, FormType, "Email", "Email", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Fax", "Fax", Equate.fmtString, Tab2 + 4750, VertPosn, 0
        
         VertPosn = 3580
         Field99Add TaxYear, FormType, "PayerFederalID", "FpayerFederalID", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "SSN", "SSN", Equate.fmtString, 2500, VertPosn, 0
         Field99Add TaxYear, FormType, "NumForms", "NumForms", Equate.fmtString, 5110, VertPosn, 0
         Field99Add TaxYear, FormType, "FWT", "FWT", Equate.fmtAmount, 6450, VertPosn, 0
         Field99Add TaxYear, FormType, "TotalAmt", "TotalAmt", Equate.fmtAmount, 8800, VertPosn, 0
         
         VertPosn = 3880
         Field99Add TaxYear, FormType, "Final", "Final", Equate.fmtString, 10400, VertPosn, 0
         
         VertPosn = 7410
         Field99Add TaxYear, FormType, "Title", "Title", Equate.fmtString, 6880, VertPosn, 0
         Field99Add TaxYear, FormType, "Date", "Date", Equate.fmtString, 9680, VertPosn, 0
         
         VertPosn = 4680
         Field99Add TaxYear, FormType, "DivX", "DivX", Equate.fmtString, 8200, VertPosn, 0
         Field99Add TaxYear, FormType, "IntX", "IntX", Equate.fmtString, 9600, VertPosn, 0
         
         VertPosn = 5670
         Field99Add TaxYear, FormType, "MiscX", "MiscX", Equate.fmtString, 1190, VertPosn, 0
         Field99Add TaxYear, FormType, "NECX", "NECX", Equate.fmtString, 1890, VertPosn, 0
         Field99Add TaxYear, FormType, "RX", "RX", Equate.fmtString, 5390, VertPosn, 0
    
    End If

End Sub



Public Sub Create2016Forms(ByVal jFormType As String)
    
Dim TaxYear As Long
Dim HorzPosn1, HorzPosn2, HorzPosn3, Tab1, Tab2 As Integer
Dim VertSpacing, VertPosn As Integer
Dim FormID As Long
Dim FormType As String

    TaxYear = 2016

    Form99.OpenRS
    Form99.Clear
    
    ' ================================================================================================
    ' 1096
    
    If jFormType = "1096" Then
    
         ClearFormRecs "1096", TaxYear
         
         Form99.FormType = "1096"
         Form99.TaxYear = TaxYear
         Form99.FormsPerPg = 1
         Form99.FormVert1 = 1000
         Form99.FormVert2 = 8950
         Form99.Save (Equate.RecAdd)
         FormID = Form99.FormID
         
         ' TaxYear, FormType,  BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
         Tab1 = 920
         Tab2 = 500
         FieldOrderNum = 0
         FormType = "1096"
         
         VertPosn = 850
         Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1450
         Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1700
         Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2150
         Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2595
         Field99Add TaxYear, FormType, "Contact", "Contact", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Phone", "Phone", Equate.fmtString, Tab2 + 4750, VertPosn, 0
             
         VertPosn = 3110
         Field99Add TaxYear, FormType, "Email", "Email", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Fax", "Fax", Equate.fmtString, Tab2 + 4750, VertPosn, 0
        
         VertPosn = 3580
         Field99Add TaxYear, FormType, "PayerFederalID", "FpayerFederalID", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "SSN", "SSN", Equate.fmtString, 2500, VertPosn, 0
         Field99Add TaxYear, FormType, "NumForms", "NumForms", Equate.fmtString, 5110, VertPosn, 0
         Field99Add TaxYear, FormType, "FWT", "FWT", Equate.fmtAmount, 6450, VertPosn, 0
         Field99Add TaxYear, FormType, "TotalAmt", "TotalAmt", Equate.fmtAmount, 8800, VertPosn, 0
         
         VertPosn = 3880
         Field99Add TaxYear, FormType, "Final", "Final", Equate.fmtString, 10400, VertPosn, 0
         
         VertPosn = 7410
         Field99Add TaxYear, FormType, "Title", "Title", Equate.fmtString, 6880, VertPosn, 0
         Field99Add TaxYear, FormType, "Date", "Date", Equate.fmtString, 9680, VertPosn, 0
         
         VertPosn = 4680
         Field99Add TaxYear, FormType, "DivX", "DivX", Equate.fmtString, 8200, VertPosn, 0
         Field99Add TaxYear, FormType, "IntX", "IntX", Equate.fmtString, 9600, VertPosn, 0
         
         VertPosn = 5670
         Field99Add TaxYear, FormType, "MiscX", "MiscX", Equate.fmtString, 1190, VertPosn, 0
         Field99Add TaxYear, FormType, "RX", "RX", Equate.fmtString, 4690, VertPosn, 0
    
    End If

End Sub

Public Sub Create2015Forms(ByVal jFormType As String)

Dim TaxYear As Long
Dim HorzPosn1, HorzPosn2, HorzPosn3, Tab1, Tab2 As Integer
Dim VertSpacing, VertPosn As Integer
Dim FormID As Long
Dim FormType As String

    TaxYear = 2015

    Form99.OpenRS
    Form99.Clear
    
    ' ================================================================================================
    ' 1099INT
    If jFormType = "All" Or jFormType = "INT" Then
    
        ClearFormRecs "INT", TaxYear
            
        Form99.FormType = "INT"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 880
        Form99.FormVert2 = 8850
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "INT"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "RTN", "Payer RTN", Equate.fmtString, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Int Inc", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "2", "Early Withdrw", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "3", "Int on US Svgs", Equate.fmtAmount, HorzPosn1, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "5", "Invest Exp", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "6", "Foreign Tax Paid", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "8", "Tax Ex Int", Equate.fmtAmount, HorzPosn1, VertPosn, 9
        Field99Add TaxYear, FormType, "9", "Spec Private Activity", Equate.fmtAmount, HorzPosn2, VertPosn, 10
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "10", "Market Discount", Equate.fmtAmount, HorzPosn1, VertPosn, 11
        Field99Add TaxYear, FormType, "11", "Bond Premium", Equate.fmtAmount, HorzPosn2, VertPosn, 12
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "12", "---", Equate.fmtAmount, HorzPosn1, VertPosn, 13
        Field99Add TaxYear, FormType, "13", "Bond Prem on tax-exempt bond", Equate.fmtAmount, HorzPosn2, VertPosn, 14
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "15a", "State", Equate.fmtString, 7400, VertPosn, 15
        Field99Add TaxYear, FormType, "16a", "State ID No.", Equate.fmtString, 8000, VertPosn, 16
        Field99Add TaxYear, FormType, "17a", "State Tax WH", Equate.fmtAmount, 9300, VertPosn, 17
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        
        Field99Add TaxYear, FormType, "15b", "State", Equate.fmtString, 7400, VertPosn, 18
        Field99Add TaxYear, FormType, "16b", "State ID No.", Equate.fmtString, 8000, VertPosn, 19
        Field99Add TaxYear, FormType, "17b", "State Tax WH", Equate.fmtAmount, 9300, VertPosn, 20
    
    End If
    
    ' ================================================================================================
    ' 1099R
    If jFormType = "All" Or jFormType = "R" Then
    
        ClearFormRecs "R", TaxYear
            
        Form99.FormType = "R"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 1000
        Form99.FormVert2 = 8950
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "R"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Gross Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "2a", "Taxable Amt", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 200
        Field99Add TaxYear, FormType, "2b1", "Tax Amt Not Determined", Equate.fmtString, HorzPosn1 + 1200, VertPosn, 3
        Field99Add TaxYear, FormType, "2b2", "Total Distr", Equate.fmtString, HorzPosn2 + 1450, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 100
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "3", "Cap Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 4
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "5", "Emp Contr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "6", "Net Unrealized", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "7", "Disc Codes", Equate.fmtString, HorzPosn1, VertPosn, 9
        Field99Add TaxYear, FormType, "8", "Other", Equate.fmtAmount, HorzPosn2 - 270, VertPosn, 10
        Field99Add TaxYear, FormType, "8P", "Box8 Pct", Equate.fmtString, HorzPosn2 + 1350, VertPosn, 11
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "9a", "Pct of Total Distr", Equate.fmtString, HorzPosn1 + 850, VertPosn, 12
        Field99Add TaxYear, FormType, "9b", "Tot Emp Contr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
        
        HorzPosn3 = HorzPosn2 + 1780
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "12a", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 13
        Field99Add TaxYear, FormType, "13a", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 14
        Field99Add TaxYear, FormType, "14a", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 15
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "10", "Amt Alloc IRR", Equate.fmtAmount, Tab1 + 300, VertPosn, 16
        Field99Add TaxYear, FormType, "11", "1st Year Roth", Equate.fmtString, Tab1 + 2400, VertPosn, 17
        Field99Add TaxYear, FormType, "12b", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 18
        Field99Add TaxYear, FormType, "13b", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 19
        Field99Add TaxYear, FormType, "14b", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 20
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "15a", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 21
        Field99Add TaxYear, FormType, "16a", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 22
        Field99Add TaxYear, FormType, "17a", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 23
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "15b", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 24
        Field99Add TaxYear, FormType, "16b", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 25
        Field99Add TaxYear, FormType, "17b", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 26
    
    End If
    
    ' ================================================================================================
    ' 1099DIV
    If jFormType = "All" Or jFormType = "DIV" Then
    
        ClearFormRecs "DIV", TaxYear
            
        Form99.FormType = "DIV"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 950
        Form99.FormVert2 = 8900
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "DIV"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1a", "Ordinary Dividends", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1b", "Qualified Dividens", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 190
        Field99Add TaxYear, FormType, "2a", "Total Cap G Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        Field99Add TaxYear, FormType, "2b", "Unrecap Sec 1250", Equate.fmtAmount, HorzPosn2, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2)
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "2c", "Sec 1202 Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "2d", "Collect 28% Gain", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 2 + 100
        Field99Add TaxYear, FormType, "3", "Nondiv Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "5", "Invest Expense", Equate.fmtAmount, HorzPosn2, VertPosn, 9
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "6", "Foreign Tax", Equate.fmtAmount, HorzPosn1, VertPosn, 10
        Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 11
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "8", "Cash Liq Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 12
        Field99Add TaxYear, FormType, "9", "NonCash Liq Distr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "10", "Exept Int Div", Equate.fmtAmount, HorzPosn1, VertPosn, 14
        Field99Add TaxYear, FormType, "11", "Spec Private Activity", Equate.fmtAmount, HorzPosn2, VertPosn, 15
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "12a", "State", Equate.fmtString, 5600, VertPosn, 16
        Field99Add TaxYear, FormType, "13a", "State ID No.", Equate.fmtString, 6200, VertPosn, 17
        Field99Add TaxYear, FormType, "14a", "State Tax WH", Equate.fmtAmount, HorzPosn2, VertPosn, 18
        
        VertPosn = VertPosn + VertSpacing - 50
        Field99Add TaxYear, FormType, "12b", "State", Equate.fmtString, 5600, VertPosn, 19
        Field99Add TaxYear, FormType, "13b", "State ID No.", Equate.fmtString, 6200, VertPosn, 20
        Field99Add TaxYear, FormType, "14b", "State Tax WH", Equate.fmtAmount, HorzPosn2, VertPosn, 21
        
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 14
    
    End If
    
    ' ================================================================================================
    ' 1099MISC
    
    If jFormType = "All" Or jFormType = "MISC" Then
    
        ClearFormRecs "MISC", TaxYear
        
        Form99.FormType = "MISC"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 1000
        Form99.FormVert2 = 8930
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "MISC"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Rents", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "2", "Royalties", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 220
        Field99Add TaxYear, FormType, "3", "Other Income", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 50
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "5", "Fishing boat proceeds", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "6", "Medical and health care payments", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "7", "Nonemployee Compensation", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "8", "Sub pmts in lieu of div or int", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "9", "Payer made direct sales", Equate.fmtString, HorzPosn1 + 1430, VertPosn, 9
        Field99Add TaxYear, FormType, "10", "Corp insurance proceeds", Equate.fmtAmount, HorzPosn2, VertPosn, 10
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "11", "Foreign Tax Paid", Equate.fmtAmount, HorzPosn1, VertPosn, 11
        Field99Add TaxYear, FormType, "12", "Foreign country or US possession", Equate.fmtString, HorzPosn2, VertPosn, 12
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "13", "Excess golden parachute", Equate.fmtAmount, HorzPosn1, VertPosn, 13
        Field99Add TaxYear, FormType, "14", "Gross proceeds to attny", Equate.fmtAmount, HorzPosn2, VertPosn, 14
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "16a", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 15
        Field99Add TaxYear, FormType, "17a", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 16
        Field99Add TaxYear, FormType, "18a", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 17
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "15a", "Sec 409A deferrals", Equate.fmtAmount, Tab1 + 300, VertPosn, 18
        Field99Add TaxYear, FormType, "15b", "Sec 409A income", Equate.fmtAmount, Tab1 + 2900, VertPosn, 19
        Field99Add TaxYear, FormType, "16b", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 20
        Field99Add TaxYear, FormType, "17b", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 21
        Field99Add TaxYear, FormType, "18b", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 22
    
    End If
    
    ' ================================================================================================
    ' 1096
    
    If jFormType = "All" Or jFormType = "1096" Then
    
         ClearFormRecs "1096", TaxYear
         
         Form99.FormType = "1096"
         Form99.TaxYear = TaxYear
         Form99.FormsPerPg = 1
         Form99.FormVert1 = 1000
         Form99.FormVert2 = 8950
         Form99.Save (Equate.RecAdd)
         FormID = Form99.FormID
         
         ' TaxYear, FormType,  BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
         Tab1 = 920
         Tab2 = 500
         FieldOrderNum = 0
         FormType = "1096"
         
         VertPosn = 850
         Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1450
         Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1700
         Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2150
         Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2595
         Field99Add TaxYear, FormType, "Contact", "Contact", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Phone", "Phone", Equate.fmtString, Tab2 + 4750, VertPosn, 0
             
         VertPosn = 3110
         Field99Add TaxYear, FormType, "Email", "Email", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Fax", "Fax", Equate.fmtString, Tab2 + 4750, VertPosn, 0
        
         VertPosn = 3580
         Field99Add TaxYear, FormType, "PayerFederalID", "FpayerFederalID", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "SSN", "SSN", Equate.fmtString, 2500, VertPosn, 0
         Field99Add TaxYear, FormType, "NumForms", "NumForms", Equate.fmtString, 5110, VertPosn, 0
         Field99Add TaxYear, FormType, "FWT", "FWT", Equate.fmtAmount, 6450, VertPosn, 0
         Field99Add TaxYear, FormType, "TotalAmt", "TotalAmt", Equate.fmtAmount, 8800, VertPosn, 0
         
         VertPosn = 3880
         Field99Add TaxYear, FormType, "Final", "Final", Equate.fmtString, 10400, VertPosn, 0
         
         VertPosn = 7410
         Field99Add TaxYear, FormType, "Title", "Title", Equate.fmtString, 6800, VertPosn, 0
         Field99Add TaxYear, FormType, "Date", "Date", Equate.fmtString, 9870, VertPosn, 0
         
         VertPosn = 4680
         Field99Add TaxYear, FormType, "DivX", "DivX", Equate.fmtString, 8570, VertPosn, 0
         Field99Add TaxYear, FormType, "IntX", "IntX", Equate.fmtString, 10220, VertPosn, 0
         
         VertPosn = 5670
         Field99Add TaxYear, FormType, "MiscX", "MiscX", Equate.fmtString, 1870, VertPosn, 0
         Field99Add TaxYear, FormType, "RX", "RX", Equate.fmtString, 4720, VertPosn, 0
    
    End If

End Sub

Public Sub Create2013Forms(ByVal jFormType As String)

Dim TaxYear As Long
Dim HorzPosn1, HorzPosn2, HorzPosn3, Tab1, Tab2 As Integer
Dim VertSpacing, VertPosn As Integer
Dim FormID As Long
Dim FormType As String

    TaxYear = 2013

    Form99.OpenRS
    Form99.Clear
    
    ' ================================================================================================
    ' 1099INT
    If jFormType = "All" Or jFormType = "INT" Then
    
        ClearFormRecs "INT", TaxYear
            
        Form99.FormType = "INT"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 880
        Form99.FormVert2 = 8850
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "INT"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "RTN", "Payer RTN", Equate.fmtString, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Int Inc", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "2", "Early Withdrw", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "3", "Int on US Svgs", Equate.fmtAmount, HorzPosn1, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "5", "Invest Exp", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "6", "Foreign Tax Paid", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "8", "Tax Ex Int", Equate.fmtAmount, HorzPosn1, VertPosn, 9
        Field99Add TaxYear, FormType, "9", "Spec Private Activity", Equate.fmtAmount, HorzPosn2, VertPosn, 10
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "11a", "State", Equate.fmtString, 7400, VertPosn, 11
        Field99Add TaxYear, FormType, "12a", "State ID No.", Equate.fmtString, 8000, VertPosn, 12
        Field99Add TaxYear, FormType, "13a", "State Tax WH", Equate.fmtAmount, 9300, VertPosn, 13
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "10", "Tax-exempt Bond", Equate.fmtString, HorzPosn1, VertPosn, 14
        
        Field99Add TaxYear, FormType, "11b", "State", Equate.fmtString, 7400, VertPosn, 15
        Field99Add TaxYear, FormType, "12b", "State ID No.", Equate.fmtString, 8000, VertPosn, 16
        Field99Add TaxYear, FormType, "13b", "State Tax WH", Equate.fmtAmount, 9300, VertPosn, 17
    
    End If
    
    ' ================================================================================================
    ' 1099R
    If jFormType = "All" Or jFormType = "R" Then
    
        ClearFormRecs "R", TaxYear
            
        Form99.FormType = "R"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 1000
        Form99.FormVert2 = 8950
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "R"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Gross Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "2a", "Taxable Amt", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 200
        Field99Add TaxYear, FormType, "2b1", "Tax Amt Not Determined", Equate.fmtString, HorzPosn1 + 1200, VertPosn, 3
        Field99Add TaxYear, FormType, "2b2", "Total Distr", Equate.fmtString, HorzPosn2 + 1450, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 100
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "3", "Cap Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 4
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "5", "Emp Contr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "6", "Net Unrealized", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "7", "Disc Codes", Equate.fmtString, HorzPosn1, VertPosn, 9
        Field99Add TaxYear, FormType, "8", "Other", Equate.fmtAmount, HorzPosn2 - 270, VertPosn, 10
        Field99Add TaxYear, FormType, "8P", "Box8 Pct", Equate.fmtString, HorzPosn2 + 1350, VertPosn, 11
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "9a", "Pct of Total Distr", Equate.fmtString, HorzPosn1 + 850, VertPosn, 12
        Field99Add TaxYear, FormType, "9b", "Tot Emp Contr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
        
        HorzPosn3 = HorzPosn2 + 1780
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "12a", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 13
        Field99Add TaxYear, FormType, "13a", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 14
        Field99Add TaxYear, FormType, "14a", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 15
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "10", "Amt Alloc IRR", Equate.fmtAmount, Tab1 + 300, VertPosn, 16
        Field99Add TaxYear, FormType, "11", "1st Year Roth", Equate.fmtString, Tab1 + 2400, VertPosn, 17
        Field99Add TaxYear, FormType, "12b", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 18
        Field99Add TaxYear, FormType, "13b", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 19
        Field99Add TaxYear, FormType, "14b", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 20
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "15a", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 21
        Field99Add TaxYear, FormType, "16a", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 22
        Field99Add TaxYear, FormType, "17a", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 23
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "15b", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 24
        Field99Add TaxYear, FormType, "16b", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 25
        Field99Add TaxYear, FormType, "17b", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 26
    
    End If
    
    ' ================================================================================================
    ' 1099DIV
    If jFormType = "All" Or jFormType = "DIV" Then
    
        ClearFormRecs "DIV", TaxYear
            
        Form99.FormType = "DIV"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 950
        Form99.FormVert2 = 8900
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "DIV"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1a", "Ordinary Dividends", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1b", "Qualified Dividens", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 190
        Field99Add TaxYear, FormType, "2a", "Total Cap G Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        Field99Add TaxYear, FormType, "2b", "Unrecap Sec 1250", Equate.fmtAmount, HorzPosn2, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2)
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "2c", "Sec 1202 Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "2d", "Collect 28% Gain", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 2 + 100
        Field99Add TaxYear, FormType, "3", "Nondiv Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "5", "Invest Expense", Equate.fmtAmount, HorzPosn2, VertPosn, 9
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "6", "Foreign Tax", Equate.fmtAmount, HorzPosn1, VertPosn, 10
        Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 11
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "8", "Cash Liq Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 12
        Field99Add TaxYear, FormType, "9", "NonCash Liq Distr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "10", "Exept Int Div", Equate.fmtAmount, HorzPosn1, VertPosn, 14
        Field99Add TaxYear, FormType, "11", "Spec Private Activity", Equate.fmtAmount, HorzPosn2, VertPosn, 15
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "12a", "State", Equate.fmtString, 5600, VertPosn, 16
        Field99Add TaxYear, FormType, "13a", "State ID No.", Equate.fmtString, 6200, VertPosn, 17
        Field99Add TaxYear, FormType, "14a", "State Tax WH", Equate.fmtAmount, HorzPosn2, VertPosn, 18
        
        VertPosn = VertPosn + VertSpacing - 50
        Field99Add TaxYear, FormType, "12b", "State", Equate.fmtString, 5600, VertPosn, 19
        Field99Add TaxYear, FormType, "13b", "State ID No.", Equate.fmtString, 6200, VertPosn, 20
        Field99Add TaxYear, FormType, "14b", "State Tax WH", Equate.fmtAmount, HorzPosn2, VertPosn, 21
        
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 14
    
    End If
    
    ' ================================================================================================
    ' 1099MISC
    
    If jFormType = "All" Or jFormType = "MISC" Then
    
        ClearFormRecs "MISC", TaxYear
        
        Form99.FormType = "MISC"
        Form99.TaxYear = TaxYear
        Form99.FormsPerPg = 2
        Form99.FormVert1 = 1000
        Form99.FormVert2 = 8930
        Form99.Save (Equate.RecAdd)
        FormID = Form99.FormID
        
        ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
        Tab1 = 500          ' demographic fields from left margin
        HorzPosn1 = 5650
        HorzPosn2 = 7400
        VertSpacing = 240
        VertPosn = 0
        FieldOrderNum = 0
        FormType = "MISC"
        
        Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "1", "Rents", Equate.fmtAmount, HorzPosn1, VertPosn, 1
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "2", "Royalties", Equate.fmtAmount, HorzPosn1, VertPosn, 2
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing * 2 - 220
        Field99Add TaxYear, FormType, "3", "Other Income", Equate.fmtAmount, HorzPosn1, VertPosn, 3
        Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 4
        
        VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 50
        Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
        Field99Add TaxYear, FormType, "5", "Fishing boat proceeds", Equate.fmtAmount, HorzPosn1, VertPosn, 5
        Field99Add TaxYear, FormType, "6", "Medical and health care payments", Equate.fmtAmount, HorzPosn2, VertPosn, 6
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "7", "Nonemployee Compensation", Equate.fmtAmount, HorzPosn1, VertPosn, 7
        Field99Add TaxYear, FormType, "8", "Sub pmts in lieu of div or int", Equate.fmtAmount, HorzPosn2, VertPosn, 8
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
        
        VertPosn = VertPosn + VertSpacing
        Field99Add TaxYear, FormType, "9", "Payer made direct sales", Equate.fmtString, HorzPosn1 + 1430, VertPosn, 9
        Field99Add TaxYear, FormType, "10", "Corp insurance proceeds", Equate.fmtAmount, HorzPosn2, VertPosn, 10
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "11", "Foreign Tax Paid", Equate.fmtAmount, HorzPosn1, VertPosn, 11
        Field99Add TaxYear, FormType, "12", "Foreign country or US possession", Equate.fmtString, HorzPosn2, VertPosn, 12
        
        VertPosn = VertPosn + VertSpacing * 3
        Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
        Field99Add TaxYear, FormType, "13", "Excess golden parachute", Equate.fmtAmount, HorzPosn1, VertPosn, 13
        Field99Add TaxYear, FormType, "14", "Gross proceeds to attny", Equate.fmtAmount, HorzPosn2, VertPosn, 14
        
        VertPosn = VertPosn + VertSpacing * 2
        Field99Add TaxYear, FormType, "16a", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 15
        Field99Add TaxYear, FormType, "17a", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 16
        Field99Add TaxYear, FormType, "18a", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 17
        
        VertPosn = VertPosn + VertSpacing * 1
        Field99Add TaxYear, FormType, "15a", "Sec 409A deferrals", Equate.fmtAmount, Tab1 + 300, VertPosn, 18
        Field99Add TaxYear, FormType, "15b", "Sec 409A income", Equate.fmtAmount, Tab1 + 2900, VertPosn, 19
        Field99Add TaxYear, FormType, "16b", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 20
        Field99Add TaxYear, FormType, "17b", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 21
        Field99Add TaxYear, FormType, "18b", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 22
    
    End If
    
    ' ================================================================================================
    ' 1096
    
    If jFormType = "All" Or jFormType = "1096" Then
    
         ClearFormRecs "1096", TaxYear
         
         Form99.FormType = "1096"
         Form99.TaxYear = TaxYear
         Form99.FormsPerPg = 1
         Form99.FormVert1 = 1000
         Form99.FormVert2 = 8950
         Form99.Save (Equate.RecAdd)
         FormID = Form99.FormID
         
         ' TaxYear, FormType,  BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
         Tab1 = 920
         Tab2 = 500
         FieldOrderNum = 0
         FormType = "1096"
         
         VertPosn = 850
         Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1450
         Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 1700
         Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2150
         Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
         
         VertPosn = 2595
         Field99Add TaxYear, FormType, "Contact", "Contact", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Phone", "Phone", Equate.fmtString, Tab2 + 4750, VertPosn, 0
             
         VertPosn = 3110
         Field99Add TaxYear, FormType, "Email", "Email", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "Fax", "Fax", Equate.fmtString, Tab2 + 4750, VertPosn, 0
        
         VertPosn = 3580
         Field99Add TaxYear, FormType, "PayerFederalID", "FpayerFederalID", Equate.fmtString, Tab2, VertPosn, 0
         Field99Add TaxYear, FormType, "SSN", "SSN", Equate.fmtString, 2500, VertPosn, 0
         Field99Add TaxYear, FormType, "NumForms", "NumForms", Equate.fmtString, 5110, VertPosn, 0
         Field99Add TaxYear, FormType, "FWT", "FWT", Equate.fmtAmount, 6450, VertPosn, 0
         Field99Add TaxYear, FormType, "TotalAmt", "TotalAmt", Equate.fmtAmount, 8800, VertPosn, 0
         
         VertPosn = 3880
         Field99Add TaxYear, FormType, "Final", "Final", Equate.fmtString, 10400, VertPosn, 0
         
         VertPosn = 7410
         Field99Add TaxYear, FormType, "Title", "Title", Equate.fmtString, 6800, VertPosn, 0
         Field99Add TaxYear, FormType, "Date", "Date", Equate.fmtString, 9870, VertPosn, 0
         
         VertPosn = 4680
         Field99Add TaxYear, FormType, "DivX", "DivX", Equate.fmtString, 6230, VertPosn, 0
         Field99Add TaxYear, FormType, "IntX", "IntX", Equate.fmtString, 7940, VertPosn, 0
         Field99Add TaxYear, FormType, "MiscX", "MiscX", Equate.fmtString, 9630, VertPosn, 0
         
         VertPosn = 5670
         Field99Add TaxYear, FormType, "RX", "RX", Equate.fmtString, 1640, VertPosn, 0
    
    End If

End Sub




Public Sub Create2011Forms()

Dim TaxYear As Long
Dim HorzPosn1, HorzPosn2, HorzPosn3, Tab1, Tab2 As Integer
Dim VertSpacing, VertPosn As Integer
Dim FormID As Long
Dim FormType As String

    TaxYear = 2011

    Form99.OpenRS
    Form99.Clear
    
    ' ================================================================================================
    ' 1099INT
    ClearFormRecs "INT", 2011
        
    Form99.FormType = "INT"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 3
    Form99.FormVert1 = 700
    Form99.FormVert2 = 6000
    Form99.FormVert3 = 11300
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "INT"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "RTN", "Payer RTN", Equate.fmtString, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1", "Int Inc", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "2", "Early Withdrw", Equate.fmtAmount, HorzPosn1, VertPosn, 3
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "3", "Int on US Svgs", Equate.fmtAmount, HorzPosn1, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "5", "Invest Exp", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "6", "Foreign Tax Paid", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "8", "Tax Ex Int", Equate.fmtAmount, HorzPosn1, VertPosn, 9
    Field99Add TaxYear, FormType, "9", "Spec Private Activity", Equate.fmtAmount, HorzPosn2, VertPosn, 10
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "10", "Tax-exempt Bond", Equate.fmtString, HorzPosn1, VertPosn, 11
    
    ' ================================================================================================
    ' 1099R
    ClearFormRecs "R", 2011
        
    Form99.FormType = "R"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 2
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "R"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1", "Gross Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "2a", "Taxable Amt", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 2 - 200
    Field99Add TaxYear, FormType, "2b1", "Tax Amt Not Determined", Equate.fmtString, HorzPosn1 + 1200, VertPosn, 3
    Field99Add TaxYear, FormType, "2b2", "Total Distr", Equate.fmtString, HorzPosn2 + 1450, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 100
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "3", "Cap Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 4
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "5", "Emp Contr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "6", "Net Unrealized", Equate.fmtAmount, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "7", "Disc Codes", Equate.fmtString, HorzPosn1, VertPosn, 9
    Field99Add TaxYear, FormType, "8", "Other", Equate.fmtAmount, HorzPosn2 - 270, VertPosn, 10
    Field99Add TaxYear, FormType, "8P", "Box8 Pct", Equate.fmtString, HorzPosn2 + 1350, VertPosn, 11
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "9a", "Pct of Total Distr", Equate.fmtString, HorzPosn1 + 850, VertPosn, 12
    Field99Add TaxYear, FormType, "9b", "Tot Emp Contr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
    
    HorzPosn3 = HorzPosn2 + 1780
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "12a", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 13
    Field99Add TaxYear, FormType, "13a", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 14
    Field99Add TaxYear, FormType, "14a", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 15
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "10", "Amt Alloc IRR", Equate.fmtAmount, Tab1 + 300, VertPosn, 16
    Field99Add TaxYear, FormType, "11", "1st Year Roth", Equate.fmtString, Tab1 + 2400, VertPosn, 17
    Field99Add TaxYear, FormType, "12b", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 18
    Field99Add TaxYear, FormType, "13b", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 19
    Field99Add TaxYear, FormType, "14b", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 20
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "15a", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 21
    Field99Add TaxYear, FormType, "16a", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 22
    Field99Add TaxYear, FormType, "17a", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 23
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "15b", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 24
    Field99Add TaxYear, FormType, "16b", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 25
    Field99Add TaxYear, FormType, "17b", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 26
    
    ' ================================================================================================
    ' 1099DIV
    ClearFormRecs "DIV", 2011
        
    Form99.FormType = "DIV"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 2
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "DIV"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1a", "Ordinary Dividends", Equate.fmtAmount, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1b", "Qualified Dividens", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 2 - 190
    Field99Add TaxYear, FormType, "2a", "Total Cap G Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 3
    Field99Add TaxYear, FormType, "2b", "Unrecap Sec 1250", Equate.fmtAmount, HorzPosn2, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2)
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "2c", "Sec 1202 Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "2d", "Collect 28% Gain", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 2 + 100
    Field99Add TaxYear, FormType, "3", "Nondiv Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "5", "Invest Expense", Equate.fmtAmount, HorzPosn2, VertPosn, 9
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "6", "Foreign Tax", Equate.fmtAmount, HorzPosn1, VertPosn, 10
    Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 11
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "8", "Cash Liq Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 12
    Field99Add TaxYear, FormType, "9", "NonCash Liq Distr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 14
    
    ' ================================================================================================
    ' 1099MISC
    
    ClearFormRecs "MISC", 2011
    
    Form99.FormType = "MISC"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 2
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "MISC"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1", "Rents", Equate.fmtAmount, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "2", "Royalties", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 2 - 100
    Field99Add TaxYear, FormType, "3", "Other Income", Equate.fmtAmount, HorzPosn1, VertPosn, 3
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2)
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "5", "Fishing boat proceeds", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "6", "Medical and health care payments", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 4
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "7", "Nonemployee Compensation", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "8", "Sub pmts in lieu of div or int", Equate.fmtAmount, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "9", "Payer made direct sales", Equate.fmtString, HorzPosn1 + 1200, VertPosn, 9
    Field99Add TaxYear, FormType, "10", "Corp insurance proceeds", Equate.fmtAmount, HorzPosn2, VertPosn, 10
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    ' 11 / 12
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "13", "Excess golden parachute", Equate.fmtAmount, HorzPosn1, VertPosn, 13
    Field99Add TaxYear, FormType, "14", "Gross proceeds to attny", Equate.fmtAmount, HorzPosn2, VertPosn, 14
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "16a", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 15
    Field99Add TaxYear, FormType, "17a", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 16
    Field99Add TaxYear, FormType, "18a", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 17
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "15a", "Sec 409A deferrals", Equate.fmtAmount, Tab1 + 300, VertPosn, 18
    Field99Add TaxYear, FormType, "15b", "Sec 409A income", Equate.fmtAmount, Tab1 + 2900, VertPosn, 19
    Field99Add TaxYear, FormType, "16b", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 20
    Field99Add TaxYear, FormType, "17b", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 21
    Field99Add TaxYear, FormType, "18b", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 22
    
    ' ================================================================================================
    ' 1096
    
    ClearFormRecs "1096", 2011
    
    Form99.FormType = "1096"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 1
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' TaxYear, FormType,  BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 920
    Tab2 = 500
    FieldOrderNum = 0
    FormType = "1096"
    
    VertPosn = 850
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 1450
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 1700
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 2150
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 2595
    Field99Add TaxYear, FormType, "Contact", "Contact", Equate.fmtString, Tab2, VertPosn, 0
    Field99Add TaxYear, FormType, "Phone", "Phone", Equate.fmtString, Tab2 + 4750, VertPosn, 0
        
    VertPosn = 3110
    Field99Add TaxYear, FormType, "Email", "Email", Equate.fmtString, Tab2, VertPosn, 0
    Field99Add TaxYear, FormType, "Fax", "Fax", Equate.fmtString, Tab2 + 4750, VertPosn, 0
   
    VertPosn = 3580
    Field99Add TaxYear, FormType, "PayerFederalID", "FpayerFederalID", Equate.fmtString, Tab2, VertPosn, 0
    Field99Add TaxYear, FormType, "SSN", "SSN", Equate.fmtString, 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "NumForms", "NumForms", Equate.fmtString, 5110, VertPosn, 0
    Field99Add TaxYear, FormType, "FWT", "FWT", Equate.fmtAmount, 6450, VertPosn, 0
    Field99Add TaxYear, FormType, "TotalAmt", "TotalAmt", Equate.fmtAmount, 8800, VertPosn, 0
    
    VertPosn = 3880
    Field99Add TaxYear, FormType, "Final", "Final", Equate.fmtString, 10400, VertPosn, 0
    
    VertPosn = 7410
    Field99Add TaxYear, FormType, "Title", "Title", Equate.fmtString, 5700, VertPosn, 0
    Field99Add TaxYear, FormType, "Date", "Date", Equate.fmtString, 9870, VertPosn, 0
    
    VertPosn = 4620
    Field99Add TaxYear, FormType, "DivX", "DivX", Equate.fmtString, 8200, VertPosn, 0
    Field99Add TaxYear, FormType, "IntX", "IntX", Equate.fmtString, 10370, VertPosn, 0
    
    VertPosn = 5610
    Field99Add TaxYear, FormType, "MiscX", "MiscX", Equate.fmtString, 2070, VertPosn, 0
    Field99Add TaxYear, FormType, "RX", "RX", Equate.fmtString, 5220, VertPosn, 0
    

End Sub

Public Sub Create2012Forms()

Dim TaxYear As Long
Dim HorzPosn1, HorzPosn2, HorzPosn3, Tab1, Tab2 As Integer
Dim VertSpacing, VertPosn As Integer
Dim FormID As Long
Dim FormType As String

    TaxYear = 2012

    Form99.OpenRS
    Form99.Clear
    
    ' ================================================================================================
    ' 1099INT
    ClearFormRecs "INT", TaxYear
        
    Form99.FormType = "INT"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 3
    Form99.FormVert1 = 700
    Form99.FormVert2 = 6000
    Form99.FormVert3 = 11300
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "INT"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "RTN", "Payer RTN", Equate.fmtString, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1", "Int Inc", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "2", "Early Withdrw", Equate.fmtAmount, HorzPosn1, VertPosn, 3
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "3", "Int on US Svgs", Equate.fmtAmount, HorzPosn1, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "5", "Invest Exp", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "6", "Foreign Tax Paid", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "8", "Tax Ex Int", Equate.fmtAmount, HorzPosn1, VertPosn, 9
    Field99Add TaxYear, FormType, "9", "Spec Private Activity", Equate.fmtAmount, HorzPosn2, VertPosn, 10
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "10", "Tax-exempt Bond", Equate.fmtString, HorzPosn1, VertPosn, 11
    
    ' ================================================================================================
    ' 1099R
    ClearFormRecs "R", TaxYear
        
    Form99.FormType = "R"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 2
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "R"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1", "Gross Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "2a", "Taxable Amt", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 2 - 200
    Field99Add TaxYear, FormType, "2b1", "Tax Amt Not Determined", Equate.fmtString, HorzPosn1 + 1200, VertPosn, 3
    Field99Add TaxYear, FormType, "2b2", "Total Distr", Equate.fmtString, HorzPosn2 + 1450, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2) + 100
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "3", "Cap Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 4
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "5", "Emp Contr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "6", "Net Unrealized", Equate.fmtAmount, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "7", "Disc Codes", Equate.fmtString, HorzPosn1, VertPosn, 9
    Field99Add TaxYear, FormType, "8", "Other", Equate.fmtAmount, HorzPosn2 - 270, VertPosn, 10
    Field99Add TaxYear, FormType, "8P", "Box8 Pct", Equate.fmtString, HorzPosn2 + 1350, VertPosn, 11
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "9a", "Pct of Total Distr", Equate.fmtString, HorzPosn1 + 850, VertPosn, 12
    Field99Add TaxYear, FormType, "9b", "Tot Emp Contr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
    
    HorzPosn3 = HorzPosn2 + 1780
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "12a", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 13
    Field99Add TaxYear, FormType, "13a", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 14
    Field99Add TaxYear, FormType, "14a", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 15
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "10", "Amt Alloc IRR", Equate.fmtAmount, Tab1 + 300, VertPosn, 16
    Field99Add TaxYear, FormType, "11", "1st Year Roth", Equate.fmtString, Tab1 + 2400, VertPosn, 17
    Field99Add TaxYear, FormType, "12b", "State Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 18
    Field99Add TaxYear, FormType, "13b", "State Num", Equate.fmtString, HorzPosn2, VertPosn, 19
    Field99Add TaxYear, FormType, "14b", "State Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 20
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "15a", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 21
    Field99Add TaxYear, FormType, "16a", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 22
    Field99Add TaxYear, FormType, "17a", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 23
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "15b", "Local Tax Withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 24
    Field99Add TaxYear, FormType, "16b", "Local Name", Equate.fmtString, HorzPosn2, VertPosn, 25
    Field99Add TaxYear, FormType, "17b", "Local Distr", Equate.fmtAmount, HorzPosn3, VertPosn, 26
    
    ' ================================================================================================
    ' 1099DIV
    ClearFormRecs "DIV", TaxYear
        
    Form99.FormType = "DIV"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 2
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "DIV"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1a", "Ordinary Dividends", Equate.fmtAmount, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1b", "Qualified Dividens", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 2 - 190
    Field99Add TaxYear, FormType, "2a", "Total Cap G Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 3
    Field99Add TaxYear, FormType, "2b", "Unrecap Sec 1250", Equate.fmtAmount, HorzPosn2, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2)
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "2c", "Sec 1202 Gain", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "2d", "Collect 28% Gain", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 2 + 100
    Field99Add TaxYear, FormType, "3", "Nondiv Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "5", "Invest Expense", Equate.fmtAmount, HorzPosn2, VertPosn, 9
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "6", "Foreign Tax", Equate.fmtAmount, HorzPosn1, VertPosn, 10
    Field99Add TaxYear, FormType, "7", "Foreign Country", Equate.fmtString, HorzPosn2, VertPosn, 11
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "8", "Cash Liq Distr", Equate.fmtAmount, HorzPosn1, VertPosn, 12
    Field99Add TaxYear, FormType, "9", "NonCash Liq Distr", Equate.fmtAmount, HorzPosn2, VertPosn, 13
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 14
    
    ' ================================================================================================
    ' 1099MISC
    
    ClearFormRecs "MISC", TaxYear
    
    Form99.FormType = "MISC"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 2
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' FormID, BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 500          ' demographic fields from left margin
    HorzPosn1 = 5650
    HorzPosn2 = 7400
    VertSpacing = 240
    VertPosn = 0
    FieldOrderNum = 0
    FormType = "MISC"
    
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "1", "Rents", Equate.fmtAmount, HorzPosn1, VertPosn, 1
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "2", "Royalties", Equate.fmtAmount, HorzPosn1, VertPosn, 2
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "Payer5", "Payer5", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = VertPosn + VertSpacing * 2 - 100
    Field99Add TaxYear, FormType, "3", "Other Income", Equate.fmtAmount, HorzPosn1, VertPosn, 3
    Field99Add TaxYear, FormType, "4", "Federal income tax withheld", Equate.fmtAmount, HorzPosn2, VertPosn, 4
    
    VertPosn = VertPosn + VertSpacing * 4 - (VertSpacing / 2)
    Field99Add TaxYear, FormType, "PayerFederalID", "PayerFederalID", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "PayeeFederalID", "PayeeFederalID", Equate.fmtString, Tab1 + 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "5", "Fishing boat proceeds", Equate.fmtAmount, HorzPosn1, VertPosn, 5
    Field99Add TaxYear, FormType, "6", "Medical and health care payments", Equate.fmtAmount, HorzPosn2, VertPosn, 6
    
    VertPosn = VertPosn + VertSpacing * 4
    Field99Add TaxYear, FormType, "PayeeName", "PayeeName", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "7", "Nonemployee Compensation", Equate.fmtAmount, HorzPosn1, VertPosn, 7
    Field99Add TaxYear, FormType, "8", "Sub pmts in lieu of div or int", Equate.fmtAmount, HorzPosn2, VertPosn, 8
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAddress", "PayeeAddress", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "9", "Payer made direct sales", Equate.fmtString, HorzPosn1 + 1200, VertPosn, 9
    Field99Add TaxYear, FormType, "10", "Corp insurance proceeds", Equate.fmtAmount, HorzPosn2, VertPosn, 10
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "PayeeCSZ", "PayeeCSZ", Equate.fmtString, Tab1, VertPosn, 0
    ' 11 / 12
    
    VertPosn = VertPosn + VertSpacing * 3
    Field99Add TaxYear, FormType, "PayeeAccountNumber", "PayeeAccountNumber", Equate.fmtString, Tab1, VertPosn, 0
    Field99Add TaxYear, FormType, "13", "Excess golden parachute", Equate.fmtAmount, HorzPosn1, VertPosn, 13
    Field99Add TaxYear, FormType, "14", "Gross proceeds to attny", Equate.fmtAmount, HorzPosn2, VertPosn, 14
    
    VertPosn = VertPosn + VertSpacing * 2
    Field99Add TaxYear, FormType, "16a", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 15
    Field99Add TaxYear, FormType, "17a", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 16
    Field99Add TaxYear, FormType, "18a", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 17
    
    VertPosn = VertPosn + VertSpacing * 1
    Field99Add TaxYear, FormType, "15a", "Sec 409A deferrals", Equate.fmtAmount, Tab1 + 300, VertPosn, 18
    Field99Add TaxYear, FormType, "15b", "Sec 409A income", Equate.fmtAmount, Tab1 + 2900, VertPosn, 19
    Field99Add TaxYear, FormType, "16b", "State tax withheld", Equate.fmtAmount, HorzPosn1, VertPosn, 20
    Field99Add TaxYear, FormType, "17b", "State/Payers state no.", Equate.fmtString, HorzPosn2, VertPosn, 21
    Field99Add TaxYear, FormType, "18b", "State tax withheld", Equate.fmtAmount, HorzPosn2 + 2100, VertPosn, 22
    
    ' ================================================================================================
    ' 1096
    
    ClearFormRecs "1096", TaxYear
    
    Form99.FormType = "1096"
    Form99.TaxYear = TaxYear
    Form99.FormsPerPg = 1
    Form99.FormVert1 = 1000
    Form99.FormVert2 = 8950
    Form99.Save (Equate.RecAdd)
    FormID = Form99.FormID
    
    ' TaxYear, FormType,  BoxName, FieldTitle, FieldFormat, HorzPosn, VertPosn, QuickEntry
    Tab1 = 920
    Tab2 = 500
    FieldOrderNum = 0
    FormType = "1096"
    
    VertPosn = 850
    Field99Add TaxYear, FormType, "Payer1", "Payer1", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 1450
    Field99Add TaxYear, FormType, "Payer2", "Payer2", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 1700
    Field99Add TaxYear, FormType, "Payer3", "Payer3", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 2150
    Field99Add TaxYear, FormType, "Payer4", "Payer4", Equate.fmtString, Tab1, VertPosn, 0
    
    VertPosn = 2595
    Field99Add TaxYear, FormType, "Contact", "Contact", Equate.fmtString, Tab2, VertPosn, 0
    Field99Add TaxYear, FormType, "Phone", "Phone", Equate.fmtString, Tab2 + 4750, VertPosn, 0
        
    VertPosn = 3110
    Field99Add TaxYear, FormType, "Email", "Email", Equate.fmtString, Tab2, VertPosn, 0
    Field99Add TaxYear, FormType, "Fax", "Fax", Equate.fmtString, Tab2 + 4750, VertPosn, 0
   
    VertPosn = 3580
    Field99Add TaxYear, FormType, "PayerFederalID", "FpayerFederalID", Equate.fmtString, Tab2, VertPosn, 0
    Field99Add TaxYear, FormType, "SSN", "SSN", Equate.fmtString, 2500, VertPosn, 0
    Field99Add TaxYear, FormType, "NumForms", "NumForms", Equate.fmtString, 5110, VertPosn, 0
    Field99Add TaxYear, FormType, "FWT", "FWT", Equate.fmtAmount, 6450, VertPosn, 0
    Field99Add TaxYear, FormType, "TotalAmt", "TotalAmt", Equate.fmtAmount, 8800, VertPosn, 0
    
    VertPosn = 3880
    Field99Add TaxYear, FormType, "Final", "Final", Equate.fmtString, 10400, VertPosn, 0
    
    VertPosn = 7410
    Field99Add TaxYear, FormType, "Title", "Title", Equate.fmtString, 5700, VertPosn, 0
    Field99Add TaxYear, FormType, "Date", "Date", Equate.fmtString, 9870, VertPosn, 0
    
    VertPosn = 4620
    Field99Add TaxYear, FormType, "DivX", "DivX", Equate.fmtString, 8200, VertPosn, 0
    Field99Add TaxYear, FormType, "IntX", "IntX", Equate.fmtString, 10370, VertPosn, 0
    
    VertPosn = 5610
    Field99Add TaxYear, FormType, "MiscX", "MiscX", Equate.fmtString, 2070, VertPosn, 0
    Field99Add TaxYear, FormType, "RX", "RX", Equate.fmtString, 5220, VertPosn, 0
    

End Sub
Private Sub Field99Add(ByVal TaxYear As Long, _
                        ByVal FormType As String, _
                        ByVal BoxName As String, _
                        ByVal FieldTitle As String, _
                        ByVal FieldFormat As Byte, _
                        ByVal HorzPosn As Long, _
                        ByVal VertPosn As Long, _
                        ByVal QuickEntry As Byte)
                        
    FieldOrderNum = FieldOrderNum + 1
    
    Field99.OpenRS
    Field99.Clear
    Field99.TaxYear = TaxYear
    Field99.FormType = FormType
    Field99.BoxName = BoxName
    Field99.FieldOrder = FieldOrderNum
    Field99.BoxName = BoxName
    Field99.FieldTitle = FieldTitle
    Field99.FieldFormat = FieldFormat
    Field99.HorzPosn = HorzPosn
    Field99.VertPosn = VertPosn
    Field99.QuickEntry = QuickEntry
    Field99.Save (Equate.RecAdd)
                        
End Sub
                        

