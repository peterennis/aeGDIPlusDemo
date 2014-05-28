Option Compare Database
Option Explicit

Public oRibbon As IRibbonUI

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