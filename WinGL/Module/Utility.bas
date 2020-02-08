Attribute VB_Name = "Utility"
Option Explicit
    
Public glCompanyName As String      ' global for current company on title bars
Public glUserName As String         ' current user name
Public glUserID As Long             ' current user ID number
Public glSuperUser As Boolean       ' Indicates user is the superuser

Public glFileName(5) As String      ' 0=Current filename
                                    ' 1-4 are the most recently used
Public glCompanyID(5) As Long

Public glLoadLast As Boolean        ' Flag to load last file on startup

'Public Sub SetGrid(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
'    gfg.FixedCols = 0                   ' see all cols selected by SQL
'    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
'    gfg.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
'    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
'    Set gfg.DataSource = grs.DataSource '
'    gfg.DataMember = grs.DataMember     '
'End Sub

Public Sub AddAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    grs.AddNew          ' Add to the recordset
    grs.Update          ' Record (save to file)
    grs.MoveLast        ' Move to the last record in the record set
    
    gfg.DataRefresh     ' Update the grid data
    gfg.Col = 0         ' Go to the first column
    gfg.SetFocus        ' Move from add button to grid

End Sub
Public Sub DelAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
' Public Sub DelAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid, ByVal Number As Long)
    grs.Delete
    grs.Update          ' Record (save to file)
'    grs.MoveLast        ' Move to the last record in the record set
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


