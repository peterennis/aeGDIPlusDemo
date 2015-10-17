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

Public Sub LoadRibbon(ByVal ribbon As IRibbonUI)
    On Error GoTo 0
    Set oRibbon = ribbon
End Sub

Public Function SetImage(ByVal rcontrol As IRibbonControl, ByRef pic As Variant) As Boolean
    On Error GoTo 0
    Set pixClass = New aeGDayClass
    Select Case rcontrol.ID
        Case "btn0"
            Set pic = pixClass.aeAttachmentToPicture(IMAGE_TABLE_NAME, "Image", "aeladdin logo 256x256.png")
        Case "btn1"
            Set pic = pixClass.aeAttachmentToPicture(IMAGE_TABLE_NAME, "Image", "photo_sceneryA32.png")
        Case "btn2"
            Set pic = pixClass.aeAttachmentToPicture(IMAGE_TABLE_NAME, "Image", "gear_refresh32.png")
        Case Else
            MsgBox "Bad SetImage Case!"
    End Select
End Function

Public Function BtnAction(ByVal rcontrol As IRibbonControl) As Boolean
    On Error GoTo 0
    Select Case rcontrol.ID
        Case "btn0"
            MsgBox "Language Selection Form", vbInformation, "GDay!"
            'DoCmd.OpenForm "frmLanguage"
            'DoEvents
            'Forms("frmLanguage").SetFocus
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

Public Sub SetLanguage(ByVal rcontrol As IRibbonControl, ByRef label)
    On Error GoTo 0
    ' Callback label
    Dim strLang As String
    Debug.Print "SetLanguage"
    Select Case rcontrol.ID
        Case "tab1"
            strLang = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 1")      ' "GDIPlus Ribbon Demo"
            label = strLang
            Debug.Print , rcontrol.ID, strLang
        Case "grp1"
            strLang = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 2")      ' "Forms"
            label = strLang
            Debug.Print , rcontrol.ID, strLang
        Case "btn0"
            strLang = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 5")      ' "Magic"
            label = strLang
            Debug.Print , rcontrol.ID, strLang
        Case "btn1"
            strLang = DLookup(aeLANG, LANGUAGE_TABLE_NAME, "LangId = 3")      ' "???"
            label = strLang
            Debug.Print , rcontrol.ID, strLang
        Case "btn2"
        Case Else
            MsgBox "Bad Language!"
    End Select
End Sub