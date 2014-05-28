Option Compare Database
Option Explicit

Dim oRibbon As IRibbonUI

Sub fuLoad(ribbon As IRibbonUI)
    Set oRibbon = ribbon
End Sub

Function fuImage(rcontrol As IRibbonControl, ByRef pic)
    Select Case rcontrol.ID
        Case "btn1"
            Set pic = AttachmentToPicture("tblImages", "Image", "photo_sceneryA32.png")
        Case "btn2"
            Set pic = AttachmentToPicture("tblImages", "Image", "gear_refresh32.png")
        Case Else
            MsgBox "Bad fuImage Case!"
    End Select
End Function

Function fuBtnAction(rcontrol As IRibbonControl)
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