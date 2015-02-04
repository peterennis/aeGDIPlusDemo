Option Compare Database
Option Explicit

Public oRibbon As IRibbonUI

Private Const aeLANG As String = "DE"
'Private Const aeLANG As String = "EN"
'

Public Sub fuLoad(ByVal ribbon As IRibbonUI)
    On Error GoTo 0
    Set oRibbon = ribbon
End Sub

Public Function fuImage(ByVal rcontrol As IRibbonControl, ByRef pic As Variant) As Boolean
    On Error GoTo 0
    Select Case rcontrol.ID
        Case "btn1"
            Set pic = AttachmentToPicture("tblImages", "Image", "photo_sceneryA32.png")
        Case "btn2"
            Set pic = AttachmentToPicture("tblImages", "Image", "gear_refresh32.png")
        Case Else
            MsgBox "Bad fuImage Case!"
    End Select
End Function

Public Function fuBtnAction(ByVal rcontrol As IRibbonControl) As Boolean
    On Error GoTo 0
    Select Case rcontrol.ID
        Case "btn1"
            DoCmd.OpenForm "frmImages"
            DoEvents
            Forms("frmImages").SetFocus
        Case "btn2"
            DoCmd.OpenForm "frmImagesOLE"
            DoEvents
            Forms("frmImagesOLE").SetFocus
        Case Else
            MsgBox "Bad fuBtnAction Case!"
    End Select
End Function

Public Sub fuLang(ByVal rcontrol As IRibbonControl, ByRef label)
    On Error GoTo 0
    ' Callback label
    Select Case rcontrol.ID
        Case "tab1"
            label = DLookup(aeLANG, "tblLanguage", "LangId = 1")      '"GDIPlus 2013 Ribbon Demo"
        Case "grp1"
            label = DLookup(aeLANG, "tblLanguage", "LangId = 2")      '"Forms"
        Case "btn0"
        Case "btn1"
            label = DLookup(aeLANG, "tblLanguage", "LangId = 3")      '"???"
        Case "btn2"
        Case Else
            MsgBox "Bad Language!"
    End Select
End Sub