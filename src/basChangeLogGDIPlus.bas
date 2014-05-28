Option Compare Database
Option Explicit

Public Const gstrDATE_GDIPlus As String = "May 28, 2014"
Public Const gstrVERSION_GDIPlus As String = "0.0.4"
Public Const gstrPROJECT_GDIPlus As String = "GDIPlusDemo"
Public Const gblnTEST_GDIPlus As Boolean = False

Public Sub GDIPlus_Export()

    Dim THE_SOURCE_FOLDER As String
    Dim THE_XML_FOLDER As String
    
    THE_SOURCE_FOLDER = "C:\ae\aeGDIPlusDemo\src\"
    THE_XML_FOLDER = "C:\ae\aeGDIPlusDemo\src\xml\"

    On Error GoTo PROC_ERR
    'aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER
    aegitClassTest varDebug:="Debugit", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GDIPlus_Export"
    Resume Next

End Sub

'=============================================================================================================================
' Tasks:
' %005 -
' %004 -
' %003 -
' %002 -
' %001 -
' Issues:
' #005 -
' #004 -
' #003 -
' #002 -
' #001 -
'=============================================================================================================================


'20140523 - v003 - Bump, fix Project Name
    ' Create tblLanguage, show USysRibbons table
    ' Rename modules
'20140523 - v002 - First commit
    ' Load aegit exp and imp classes
    ' Configure system for export