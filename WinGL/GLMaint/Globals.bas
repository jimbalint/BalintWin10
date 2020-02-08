Attribute VB_Name = "Globals"
Option Explicit

    Public cn As New ADODB.Connection
    
    Public glCompanyName As String      ' global for current company on title bars
    Public glUserName As String         ' current user name
    Public glUserID As Long             ' current user ID number
    Public glSuperUser As Boolean       ' Indicates user is the superuser
    
    Public glFileName(5) As String      ' 0=Current filename
                                        ' 1-4 are the most recently used
    Public glLoadLast As Boolean        ' Flag to load last file on startup
    
    
    
