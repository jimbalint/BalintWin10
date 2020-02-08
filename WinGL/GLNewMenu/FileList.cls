VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FileName As String                   ' Collection class  of filenames
Private mcol As New Collection              ' maintained by the class

Public Sub LoadCmb(ByRef cmb As ComboBox)   ' load any combo box
    If mcol.Count = 0 Then Exit Sub         ' no adds if none on list
    cmb.Clear                               ' clear current combo items
    For index = mcol.Count To 1 Step -1     ' loads last on top
        cmb.AddItem mcol.Item(index).FileName
    Next index
    cmb.Text = mcol(mcol.Count).FileName    ' set box text to last loaded
End Sub

Public Function Records() As Long
    Records = mcol.Count            ' works like a queue function in C5.5
End Function

Public Function Item(ByVal index As Variant) As FileList
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)     ' Use Tools\Procedure Attributes\Advanced
End Function                        ' To set this to default ALWAYS in a cc

Public Function Free()              ' works like free(queue) in C5.5
    Dim index As Variant
    For index = mcol.Count To 1 Step -1
        mcol.Remove index
    Next index
End Function

Public Sub Delete(ByVal index As Variant)   ' like delete(queue) in C5.5
    mcol.Remove index
End Sub

Public Function Add(ByVal sFileName As String)  ' Add each file user picks
    On Error GoTo MatErr
    Dim cc As New FileList
    cc.FileName = sFileName
    mcol.Add cc, cc.FileName
    Set Add = cc
MatErr:
End Function
