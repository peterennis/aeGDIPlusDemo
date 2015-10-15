Option Compare Database
Option Explicit

Public oRibbon As IRibbonUI

Private Const aeLANG As String = "DE"
'Private Const aeLANG As String = "EN"
Public Const IMAGE_TABLE_NAME As String = "tblImages"
Public Const OLE_IMAGE_TABLE_NAME As String = "tblOLE"
Private Const LANGUAGE_TABLE_NAME As String = "tblLanguage"
Private pixClass As aeGDayClass
'

Public Sub fuLoad(ByVal ribbon As IRibbonUI)
    On Error GoTo 0
    Set oRibbon = ribbon
End Sub

Public Function SetImage(ByVal rcontrol As IRibbonControl, ByRef pic As Variant) As Boolean
    On Error GoTo 0
    Set pixClass = New aeGDayClass
    Select Case rcontrol.ID
        Case "btn1"
            Set pic.Handle = pixClass.aeAttachmentToPicture(IMAGE_TABLE_NAME, "Image", "photo_sceneryA32.png")
        Case "btn2"
            Set pic.Handle = pixClass.aeAttachmentToPicture(IMAGE_TABLE_NAME, "Image", "gear_refresh32.png")
        Case Else
            MsgBox "Bad SetImage Case!"
    End Select
End Function

Public Function BtnAction(ByVal rcontrol As IRibbonControl) As Boolean
    On Error GoTo 0
    Select Case rcontrol.ID
        Case "btn1"
            DoCmd.OpenForm "frmClassImages"
            DoEvents
            Forms("frmClassImages").SetFocus
        Case "btn2"
            DoCmd.OpenForm "frmClassImagesOLE"
            DoEvents
            Forms("frmClassImagesOLE").SetFocus
        Case Else
            MsgBox "Bad BtnAction Case!"
    End Select
End Function

Public Sub fuLang(ByVal rcontrol As IRibbonControl, ByRef label)
    On Error GoTo 0
    ' Callback label
    Select Case rcontrol.ID
        Case "tab1"
            label = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 1")      '"GDIPlus 2013 Ribbon Demo"
        Case "grp1"
            label = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 2")      '"Forms"
        Case "btn0"
        Case "btn1"
            label = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 3")      '"???"
        Case "btn2"
        Case Else
            MsgBox "Bad Language!"
    End Select
End Sub