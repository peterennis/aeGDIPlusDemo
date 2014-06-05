Option Compare Database
Option Explicit

Public Sub GDIPlusDemo2013_Export()

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
' %001 - GdiPlus leaks - Ref: http://blogs.msdn.com/b/dsui_team/archive/2013/04/23/debugging-a-gdi-resource-leak.aspx
' Issues:
' #005 -
' #004 -
' #003 -
' #002 -
' #001 -
'=============================================================================================================================


'20140605 - v008 - Include basDemoVersion, move constants from change log
    ' GDIPlusDemo2013_fixed.zip (v007) includes solution for drag and drop error
    ' Add reference to MSCOMCTL.OCX, compile, remove reference - fixes ActiveX
    ' registration problem in the forms.
'20140604 - v007 - Fixed button picture load error in GDIPlus from Sascha
'20140529 - v006 - Windows 8 DLL File Information - GdiPlus.dll
    ' Ref: http://www.nirsoft.net/dll_information/windows8/gdiplus_dll.html
    ' Credit: mossSOFT, Entwicklung und Beratung, Sascha Alexander Trowitzsch
    ' Use latest module from here:
    ' Ref: http://www.mosstools.de/index.php?option=com_content&view=article&id=77&Itemid=76
    ' Office 2013 Setup distributes neither ogl.dll nor gdiplus.dll. The prerequisites for O2013 are that you install it on a system already containing gdiplus 1.1. So there's no need to include it.
    ' To workaround it use the latest version of the module and alter just one line in procedure GetGDIPVersion:
    ' Change Case "11.0" 'A2003 -to- Case "11.0", "15.0" 'A2003, A2013
'20140528 - v005 - Fixes using TM VBA-Inspector
    ' Use a space before comments
    ' Office 2013 - Ref: http://www.utteraccess.com/forum/Custom-Ribbon-Icon-Ogl-t2016045.html
    ' s/ogl/GDIPlus/g will fix the demo to run in Access 2013
    ' Also changed ogl.dll to gdiplus.dll
    ' Based on work from here:
    ' Ref: http://www.activevb.de/tipps/vb6tipps/tipp0644.html
    ' Add version/date/project details to forms
    ' Use export modules only
'20140527 - v004 - Fix IsQryHidden problem with export
'20140523 - v003 - Bump, fix Project Name
    ' Create tblLanguage, show USysRibbons table
    ' Rename modules
'20140523 - v002 - First commit
    ' Load aegit exp and imp classes
    ' Configure system for export