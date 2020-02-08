VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form CompanyList 
   Caption         =   "Company Records"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "CompanyList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtLocator 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin TrueDBGrid80.TDBGrid Grid1 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7223
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Client Company"
      Columns(0).DataField=   ""
      Columns(0).DataWidth=   30
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "File Name"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Default         =   -1  'True
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "&Locator"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "CompanyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xdb As New XArrayDB
Dim order As XORDER

Private Sub cmdDelete_Click()
    
Dim Resp As Integer
Dim FName As String
Dim x As String
Dim Rw As Integer
    
Dim db As DAO.Database
    
    ' xdb col 0 = company name
    '     col 1 = file name
    '     col 2 = company ID
    
    Rw = Grid1.Bookmark
    
    Resp = MsgBox("Are you SURE you want to delete " & xdb(Grid1.Bookmark, 0) & " ???", _
                   vbExclamation + vbOKCancel + vbDefaultButton2, "DELETE COMPANY INFORMATION")
    If Resp = vbCancel Then Exit Sub
    
    Resp = MsgBox("ALL INFORMATION FOR " & xdb(Grid1.Bookmark, 0) & " WILL BE DELETED !!!", _
                   vbExclamation + vbOKCancel + vbDefaultButton2, "DELETE COMPANY INFORMATION")
    If Resp = vbCancel Then Exit Sub
    
    ' delete the file from disk
    x = xdb(Rw, 1)
    FName = Mid(App.Path, 1, 2) & Mid(x, 3, Len(x) - 2)
    Kill FName
    
    ' clear last company from user record if deleting it
    ' and disable menu items
    If xdb(Rw, 2) = User(1).LastCompany Then
       User(1).LastCompany = 0
       User(1).PutRecord (UserID)
       MainMenu.SetCompany (0)
    End If
    
    ' delete from the company file
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    x = "DELETE * FROM GLCompany WHERE ID = " & xdb(Rw, 2)
    db.Execute x
    
    ' remove from the XArray
    xdb.DeleteRows (Rw)
    
    ' update the grid
    Grid1.ReBind
    Grid1.Bookmark = 1
    
End Sub

'Private Sub cmdAdd_Click()
'    On Error GoTo glErr
'    CompanyForm.ID = 0
'    CompanyForm.Init
'    CompanyForm.Show vbModal
'    If CompanyForm.userOK = True Then
'        xdb.AppendRows
'        xdb.Value(xdb.UpperBound(1), 0) = CStr(CompanyForm.txtName)
'        xdb.Value(xdb.UpperBound(1), 1) = CStr(CompanyForm.ID)
'    End If
'    Unload CompanyForm
'    Grid1.ReBind
'    Grid1.Bookmark = xdb.UpperBound(1)
'    Grid1.SetFocus
'    Exit Sub
'glErr:
'    MsgBox Error(Err.Number)
'End Sub

'Private Sub cmdEdit_Click()
'    On Error GoTo glErr
'    CompanyForm.ID = CLng(xdb.Value(Grid1.Bookmark, 1))
'    CompanyForm.Init
'    CompanyForm.Show vbModal
'    xdb.Value(Grid1.Bookmark, 0) = CompanyForm.txtName
'    Grid1.RefetchRow
'    Unload CompanyForm
'    Grid1.SetFocus
'    Exit Sub
'glErr:
'    MsgBox Error(Err.Number)
'End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    MainMenu.SetCompany CLng(xdb.Value(Grid1.Bookmark, 2))
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo glErr
    Set xdb = xFactory.GetCompany("\balint\data\glSystem.mdb")
    Set Grid1.Array = xdb
    
    Grid1.Columns(0).Width = 3000
    Grid1.Columns(1).Width = 5000
    Dim ndx As Integer
    For ndx = xdb.LowerBound(1) To xdb.UpperBound(1)
        If curCompany = CLng(xdb.Value(ndx, 2)) Then
            Grid1.Bookmark = ndx
        End If
    Next ndx
    order = XORDER_ASCEND
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload CompanyForm
End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    Dim ColType As String
    ColType = XTYPE_STRING
    If order = XORDER_ASCEND Then
        xdb.QuickSort xdb.LowerBound(1), xdb.UpperBound(1), ColIndex, XORDER_DESCEND, ColType
        order = XORDER_DESCEND
    Else
        xdb.QuickSort xdb.LowerBound(1), xdb.UpperBound(1), ColIndex, XORDER_ASCEND, ColType
        order = XORDER_ASCEND
    End If
    Grid1.ClearSelCols
    Grid1.Refresh
End Sub

Private Sub txtLocator_LostFocus()
    txtLocator = ""
End Sub

Private Sub txtLocator_Change()
    Grid1.Bookmark = xFactory.Locator(txtLocator, xdb, 0, Grid1.Bookmark)
    Grid1.Refresh
End Sub

