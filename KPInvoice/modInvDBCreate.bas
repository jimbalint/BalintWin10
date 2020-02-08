Attribute VB_Name = "modInvDBCreate"
Option Explicit

Public Sub StockCreate(Optional ByVal ReCreate As Boolean = False)

    ' ****************************************************
    If TableExists("InvStock", cn) = True Then
        If ReCreate = True Then
            SQLString = "DROP TABLE InvStock"
            cn.Execute SQLString
        Else
            MsgBox "InvStock table already exists!", vbExclamation
            Exit Sub
        End If
    End If
    ' ****************************************************

    SQLString = "CREATE TABLE InvStock ( " & _
                "[StockID] Counter, CONSTRAINT stkIDKey PRIMARY KEY ([StockID]) ) "
    cn.Execute SQLString

    AddField "InvStock", "StockSelect", "Logical", cn
    AddField "InvStock", "QBID", "Char (50)", cn
    AddField "InvStock", "QBName", "Char (159)", cn
    AddField "InvStock", "JobID", "Long", cn
    AddField "InvStock", "Show", "Byte", cn
    AddField "InvStock", "Description", "Char (255)", cn
    AddField "InvStock", "MasterPrice", "Double", cn
    AddField "InvStock", "CustomerPrice", "Double", cn
    AddField "InvStock", "Cost", "Double", cn
    AddField "InvStock", "LastDate", "DateTime", cn
    AddField "InvStock", "Active", "Logical", cn
    AddField "InvStock", "InventoryItem", "Logical", cn

End Sub

Public Sub HeaderCreate(Optional ByVal ReCreate As Boolean = False)

    ' ****************************************************
    If TableExists("InvHeader", cn) = True Then
        If ReCreate = True Then
            SQLString = "DROP TABLE InvHeader"
            cn.Execute SQLString
        Else
            MsgBox "InvHeader table already exists!", vbExclamation
            Exit Sub
        End If
    End If
    ' ****************************************************

    SQLString = "CREATE TABLE InvHeader ( " & _
                "[HeaderID] Counter, CONSTRAINT hdrIDKey PRIMARY KEY ([HeaderID]) ) "
    cn.Execute SQLString

    AddField "InvHeader", "SoldAddr1", "Char (40)", cn
    AddField "InvHeader", "SoldAddr2", "Char (40)", cn
    AddField "InvHeader", "SoldAddr3", "Char (40)", cn
    AddField "InvHeader", "SoldAddr4", "Char (40)", cn
    AddField "InvHeader", "SoldCity", "Char (40)", cn
    AddField "InvHeader", "SoldState", "Char (2)", cn
    AddField "InvHeader", "SoldZip", "Char (10)", cn
    
    AddField "InvHeader", "ShipAddr1", "Char (40)", cn
    AddField "InvHeader", "ShipAddr2", "Char (40)", cn
    AddField "InvHeader", "ShipAddr3", "Char (40)", cn
    AddField "InvHeader", "ShipAddr4", "Char (40)", cn
    AddField "InvHeader", "ShipCity", "Char (40)", cn
    AddField "InvHeader", "ShipState", "Char (2)", cn
    AddField "InvHeader", "ShipZip", "Char (10)", cn
    
    AddField "InvHeader", "SoldJobID", "Long", cn
    AddField "InvHeader", "InvoiceNumber", "Long", cn
    AddField "InvHeader", "SaveFlag", "Byte", cn
    AddField "InvHeader", "OrderDate", "DateTime", cn
    AddField "InvHeader", "InvoiceDate", "DateTime", cn
    AddField "InvHeader", "PackageCount", "Long", cn
    AddField "InvHeader", "PalletCount", "Long", cn
    AddField "InvHeader", "ItemTotal", "Currency", cn
    AddField "InvHeader", "SalesTax", "Currency", cn
    AddField "InvHeader", "Freight", "Currency", cn
    AddField "InvHeader", "TotalAmount", "Currency", cn
    AddField "InvHeader", "PO1", "Char (50)", cn
    AddField "InvHeader", "PO2", "Char (50)", cn
    AddField "InvHeader", "TruckID1", "Long", cn
    AddField "InvHeader", "TruckID2", "Long", cn
    AddField "InvHeader", "TruckID3", "Long", cn
    AddField "InvHeader", "TrailerID1", "Long", cn
    AddField "InvHeader", "TrailerID2", "Long", cn
    AddField "InvHeader", "TrailerID3", "Long", cn
    AddField "InvHeader", "DriverID1", "Long", cn
    AddField "InvHeader", "DriverID2", "Long", cn
    AddField "InvHeader", "DriverID3", "Long", cn
    AddField "InvHeader", "ApptDate", "DateTime", cn
    AddField "InvHeader", "ApptTime", "Char (10)", cn
    AddField "InvHeader", "Terms", "Char (50)", cn
    AddField "InvHeader", "QBInvoiceID", "String (50)", cn

End Sub

Public Sub BodyCreate(Optional ByVal ReCreate As Boolean = False)

    ' ****************************************************
    If TableExists("InvBody", cn) = True Then
        If ReCreate = True Then
            SQLString = "DROP TABLE InvBody"
            cn.Execute SQLString
        Else
            MsgBox "InvBody table already exists!", vbExclamation
            Exit Sub
        End If
    End If
    ' ****************************************************

    SQLString = "CREATE TABLE InvBody ( " & _
              "[BodyID] Counter, CONSTRAINT bdyIDKey PRIMARY KEY ([BodyID]) )"
    cn.Execute SQLString

    AddField "InvBody", "HeaderID", "Long", cn
    AddField "InvBody", "LineNum", "Long", cn
    AddField "InvBody", "QtyOrdered", "Double", cn
    AddField "InvBody", "QtyShipped", "Double", cn
    AddField "InvBody", "Description", "Char (255)", cn
    AddField "InvBody", "StockID", "Long", cn
    AddField "InvBody", "Price", "Double", cn
    AddField "InvBody", "Amount", "Double", cn

End Sub

Public Sub InvGlobalCreate(Optional ByVal ReCreate As Boolean = False)

    ' ****************************************************
    If TableExists("InvGlobal", cnDes) = True Then
        If ReCreate = True Then
            SQLString = "DROP TABLE InvGlobal"
            cnDes.Execute SQLString
        Else
            MsgBox "InvGlobal table already exists!", vbExclamation
            Exit Sub
        End If
    End If
    ' ****************************************************

    SQLString = "CREATE TABLE InvGlobal ( " & _
              "[GlobalID] Counter, CONSTRAINT glbIDKey PRIMARY KEY ([GlobalID]) )"
              cnDes.Execute SQLString

    AddField "InvGlobal", "CompanyID", "Long", cnDes
    AddField "InvGlobal", "UserID", "Long", cnDes
    AddField "InvGlobal", "TypeCode", "Byte", cnDes
    AddField "InvGlobal", "Description", "Char (255)", cnDes
    AddField "InvGlobal", "Byte1", "Byte", cnDes
    AddField "InvGlobal", "Byte2", "Byte", cnDes
    AddField "InvGlobal", "Byte3", "Byte", cnDes
    AddField "InvGlobal", "Byte4", "Byte", cnDes
    AddField "InvGlobal", "Byte5", "Byte", cnDes
    AddField "InvGlobal", "Byte6", "Byte", cnDes
    AddField "InvGlobal", "Byte7", "Byte", cnDes
    AddField "InvGlobal", "Byte8", "Byte", cnDes
    AddField "InvGlobal", "Byte9", "Byte", cnDes
    AddField "InvGlobal", "Byte10", "Byte", cnDes
    AddField "InvGlobal", "Var1", "Char (50)", cnDes
    AddField "InvGlobal", "Var2", "Char (50)", cnDes
    AddField "InvGlobal", "Var3", "Char (50)", cnDes
    AddField "InvGlobal", "Var4", "Char (50)", cnDes
    AddField "InvGlobal", "Var5", "Char (50)", cnDes

End Sub


