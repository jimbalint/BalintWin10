VERSION 5.00
Begin VB.Form Form1 
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim rs As New ADODB.Recordset
   Dim rsdt As ADODB.Recordset


Private Sub Form_Load()
   Dim sTable As String
   Dim sNewTable As String
   Dim x As String
   Dim Channel As Integer
   Dim Pre As String
   Dim dt As Integer
   Dim CT As Long

   Dim cn As New ADODB.Connection
   
   cn.ConnectionString = "DSN=winclub"
   
   cn.Provider = "Microsoft.Jet.OLEDB.3.51"
   cn.ConnectionString = "\data\dlrcont.mdb"
   cn.Open
   
   rs.ActiveConnection = cn
   rs.CursorType = adOpenDynamic
   rs.LockType = adLockOptimistic
   Set rs = cn.OpenSchema(adSchemaColumns)
   
   Channel = FreeFile
   Open "\asend\DlrCont.txt" For Output As Channel
   
   Do Until rs.EOF = True
      sTable = rs!Table_Name
      
      If Left(sTable, 4) = "MSys" Then GoTo Cycle
      
      If (sTable <> sNewTable) Then
         
         x = ""
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = rs!Table_Name
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = ""
         Print #Channel, x
      
      End If
      sNewTable = sTable
      
      dt = rs!Data_Type
      
      ' LET !!!!!!!!!!!!!!!!!
      x = "Public Property Let " & rs!Column_Name & "(ByVal " & Prefix(dt) & "New as " & VType(dt) & ")"
      Print #Channel, x
      
      x = "    m" & Prefix(dt) & rs!Column_Name & " = " & Prefix(dt) & "New"
      Print #Channel, x
      
      x = "End Property"
      Print #Channel, x
      
      x = ""
      Print #Channel, x
      
      ' GET !!!!!!!!!!!!
      x = "Public Property Get " & rs!Column_Name & "() as " & VType(dt)
      Print #Channel, x
      
      x = "    " & rs!Column_Name & " = m" & Prefix(dt) & rs!Column_Name
      
      If rs!Data_Type = 129 Then
         '  & ""
         x = x & " " & Chr(38) & " " & Chr(34) & Chr(34)
      End If
      
      Print #Channel, x
      
      x = "End Property"
      Print #Channel, x
      
      x = ""
      Print #Channel, x
      Print #Channel, x
      Print #Channel, x
      
Cycle:
      rs.MoveNext
   Loop
   
   ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   
   rs.MoveFirst
   
   Do Until rs.EOF = True
      sTable = rs!Table_Name
      
      If Left(sTable, 4) = "MSys" Then GoTo Cycle2
      
      If (sTable <> sNewTable) Then
         x = ""
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = rs!Table_Name
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = ""
         Print #Channel, x
      
      End If
      sNewTable = sTable
      
      x = "rs.Fields(""" & rs!Column_Name & """) = " & rs!Column_Name
      
      If rs!Data_Type = 129 Then
         '  & ""
         x = x & " " & Chr(38) & " " & Chr(34) & Chr(34)
      End If
      
      Print #Channel, x
      
Cycle2:
      rs.MoveNext
   Loop
   
   ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   
   rs.MoveFirst
   
   Do Until rs.EOF = True
      sTable = rs!Table_Name
      
      If Left(sTable, 4) = "MSys" Then GoTo Cycle3
      
      If (sTable <> sNewTable) Then
         x = ""
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = rs!Table_Name
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = ""
         Print #Channel, x
      
      End If
      sNewTable = sTable
      
      dt = rs!Data_Type
      
      x = "Private m" & Prefix(dt) & rs!Column_Name & " as " & VType(dt)
      Print #Channel, x
      
Cycle3:
      rs.MoveNext
   Loop
   
   ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   
   rs.MoveFirst
   
   Do Until rs.EOF = True
      sTable = rs!Table_Name
      
      If Left(sTable, 4) = "MSys" Then GoTo CycleDot
      
      If (sTable <> sNewTable) Then
         x = ""
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = rs!Table_Name
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = ""
         Print #Channel, x
      
      End If
      sNewTable = sTable
      
      dt = rs!Data_Type
      
      x = "         " & _
          Chr(34) & rs!Table_Name & _
          "." & rs!Column_Name & _
          " " & Chr(34) & " " & Chr(38) & " " & Chr(95)
      Print #Channel, x
      
CycleDot:
      rs.MoveNext
   Loop
   
   ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   
   
   rs.MoveFirst
   
   Do Until rs.EOF = True
      sTable = rs!Table_Name
      
      If Left(sTable, 4) = "MSys" Then GoTo Cycle4
      
      If (sTable <> sNewTable) Then
         x = ""
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = rs!Table_Name
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = ""
         Print #Channel, x
      
      End If
      sNewTable = sTable
      
      dt = rs!Data_Type
      
      x = "     " & rs!Column_Name & " = rs!" & rs!Column_Name
      
      If rs!Data_Type = 129 Then
         '  & ""
         x = x & " " & Chr(38) & " " & Chr(34) & Chr(34)
      End If
      
      Print #Channel, x
      
Cycle4:
      rs.MoveNext
   Loop
   
   ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   
   rs.MoveFirst
   
   CT = 0
   
   Do Until rs.EOF = True
      sTable = rs!Table_Name
      
      If Left(sTable, 4) = "MSys" Then GoTo Cycle5
      
      If (sTable <> sNewTable) Then
         
         If CT <> 0 Then
            Print #Channel, x & Chr(34) & Chr(38) & " " & Chr(95)
            CT = 0
         End If
         
         If sNewTable <> "" Then
            x = "     " & Chr(34) & "FROM " & sNewTable & Chr(34)
            Print #Channel, x
         End If
         
         x = ""
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = rs!Table_Name
         Print #Channel, x
         
         x = "=================================="
         Print #Channel, x
         
         x = ""
         Print #Channel, x
      
         x = "     " & Chr(34) & "SELECT " & Chr(34) & Chr(38) & " " & Chr(95)
         Print #Channel, x
         
         x = "     " & Chr(34)
      
      End If
      sNewTable = sTable
      
      dt = rs!Data_Type
      
      x = x & rs!Table_Name & "." & rs!Column_Name & ", "
      
      CT = CT + 1
      If CT = 5 Then
         x = x & Chr(34) & Chr(38) & " " & Chr(95)
         Print #Channel, x
         x = "     " & Chr(34)
         CT = 0
      End If
      
Cycle5:
      rs.MoveNext
   Loop
   
   End

End Sub



Public Function Prefix(ByVal DType As Integer) As String
   Select Case DType
       Case 2
          Prefix = "int"
       Case 3
          Prefix = "lng"
       Case 6
          Prefix = "cur"
       Case 7
          Prefix = "dat"
       Case 11
          Prefix = "boo"
       Case 16
          Prefix = "byt"
       Case 17
          Prefix = "byt"
       Case 129
          Prefix = "txt"
       Case 131
          Prefix = "cur"
       Case 133
          Prefix = "dat"
       Case Else
          Prefix = "xxx"
   End Select
End Function


Public Function VType(ByVal DType As Integer) As String
   
   Select Case DType
       Case 2
          VType = "Integer"
       Case 3
          VType = "Long"
       Case 6
          VType = "Currency"
       Case 7
          VType = "Date"
       Case 11
          VType = "Boolean"
       Case 16
          VType = "Byte"
       Case 17
          VType = "Byte"
       Case 129
          VType = "String"
       Case 131
          VType = "Currency"
       Case 133
          VType = "Date"
       Case Else
          VType = CStr(DType)
   End Select
End Function

