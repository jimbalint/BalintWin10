Attribute VB_Name = "modGlobal"
Public Sub SetGrid(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    gfg.FixedCols = 0                   ' see all cols selected by SQL
    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    gfg.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set gfg.DataSource = grs.DataSource '
    gfg.DataMember = grs.DataMember     '

    gfg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    gfg.TabBehavior = flexTabCells                       ' tab moves between cells
    gfg.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    ' gfg.HighLight = flexHighlightNever                   ' don't select ranges

End Sub

