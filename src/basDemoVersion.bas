Option Compare Database
Option Explicit

Private Const gstrVERSION_GDIPlus As String = "0.1.6"
Private Const gstrDATE_GDIPlus As String = "October 15, 2015"
Public Const gstrPROJECT_GDIPlus As String = "GDayClass"
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