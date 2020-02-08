Attribute VB_Name = "Utility"
Option Explicit

Public Sub SetGrid(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    gfg.FixedCols = 0                   ' see all cols selected by SQL
    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    gfg.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set gfg.DataSource = grs.DataSource '
    gfg.DataMember = grs.DataMember     '
End Sub

Public Sub AddAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    grs.AddNew          ' Add to the recordset
    grs.Update          ' Record (save to file)
    grs.MoveLast        ' Move to the last record in the record set
    gfg.DataRefresh     ' Update the grid data
    gfg.Col = 0         ' Go to the first column
    gfg.SetFocus        ' Move from add button to grid
End Sub

Public Sub SetAdo(ByRef gcn As ADODB.Connection, ByRef grs As ADODB.Recordset, ByVal SQL As String)
    ' Common behavior for Recordsets
    Set grs = New ADODB.Recordset       ' set the recordset
    grs.LockType = adLockOptimistic     '
    grs.CursorType = adOpenDynamic      '
    grs.Source = SQL                    '
    Set grs.ActiveConnection = gcn      ' connection set previous to call
    grs.Open                            ' start the record
End Sub

Public Function glConnect() As Boolean
On Error GoTo ErrHandler
    If cn.State = False Then
        cn.ConnectionString = " Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & glFileName(0)
        cn.Open     ' Common connection for client tables file
    End If
    Dim mrs As ADODB.Recordset
    SetAdo cn, mrs, "select * from GLCompany"
    glCompanyName = mrs!Name
    glConnect = True
    Exit Function
ErrHandler:
    MsgBox "GL File Not Open", , "GL Menu"
    glConnect = False
End Function

