Option Compare Database
Option Explicit

Private Const gstrVERSION_GDIPlus As String = "0.0.8"
Private Const gstrDATE_GDIPlus As String = "June 5, 2014"
Public Const gstrPROJECT_GDIPlus As String = "GDIPlusDemo2013"
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = gstrVERSION_GDIPlus
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = gstrDATE_GDIPlus
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_GDIPlus
End Function