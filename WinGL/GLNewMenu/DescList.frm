VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DescList 
   Caption         =   " General Ledger Descriptions"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   Icon            =   "DescList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.ListView List1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NUMBER"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DESCRIPTION"
         Object.Width           =   14111
      EndProperty
   End
End
Attribute VB_Name = "DescList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curDesc As Long

Private Sub cmdAdd_Click()
    On Error GoTo glErr
    ndx = List1.SelectedItem.index
    DescForm.ID = 0
    DescForm.Init
    DescForm.Show vbModal
    If DescForm.userOK = True Then
        List1.ListItems.Add ndx, , DescForm.txtNumber
        List1.ListItems(ndx).SubItems(1) = DescForm.txtDescription
    End If
    Unload DescForm
    List1.ListItems(ndx).Selected = True
    List1.SetFocus
    Exit Sub
glErr:
    MsgBox Error(Err.number)
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo glErr
    DescForm.txtNumber = List1.SelectedItem
    DescForm.Init
    DescForm.Show vbModal
    If DescForm.userOK = True Then
        List1.SelectedItem.Text = DescForm.txtNumber
        List1.SelectedItem.SubItems(1) = DescForm.txtDescription
    End If
    Unload DescForm
    List1.SetFocus
    Exit Sub
glErr:
    MsgBox Error(Err.number)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo glErr
    Dim ndx As Long
    Dim cc As New ccDescriptions
    cc.GetSQL ("select * from glDescriptions order by number")
    For ndx = 1 To cc.Records
        List1.ListItems.Add ndx, , cc(ndx).number
        List1.ListItems(ndx).SubItems(1) = cc(ndx).description
    Next ndx
    Exit Sub
glErr:
End Sub
