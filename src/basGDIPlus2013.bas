Option Compare Database
Option Explicit

'-------------------------------------------------
'    Picture functions using GDIPlus-API (GDIP)   |
'-------------------------------------------------
'   (c) mossSOFT / Sascha Trowitzsch rev. 08/2009 |
'-------------------------------------------------
'    *  Office 2013 version  *                    |
'    rev. 05/2014 Peter F. Ennis                  |
'-------------------------------------------------

' Reference to library "OLE Automation" (stdole) needed!

Public Const GUID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"    ' IPicture

' User-defined types: ----------------------------------------------------------------------

Public Enum PicFileType
    pictypeBMP = 1
    pictypeGIF = 2
    pictypePNG = 3
    pictypeJPG = 4
End Enum

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type TSize
    X As Double
    Y As Double
End Type

Public Type RECT
    Bottom As Long
    Left As Long
    Right As Long
    Top As Long
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type GDIPStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    UUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type

' API-Declarations: ----------------------------------------------------------------------------

' Convert a windows bitmap to OLE-Picture :
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
' Retrieve GUID-Type from string :
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pCLSID As GUID) As Long

' Memory functions:
'''NOT USED - Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
'''NOT USED - Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'''NOT USED - Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Byte, ByVal Length As Long)

' Modules API:
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

' Timer API:
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

' OLE-Stream functions :
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As Long) As Long

' GDIPlus Flat-API declarations:
' *Remark: We use gdiplus.dll

' Initialization GDIP:
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
' Tear down GDIP:
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
' Load GDIP-Image from file :
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As Long, BITMAP As Long) As Long
' Create GDIP- graphical area from Windows-DeviceContext:
Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hdc As Long, GpGraphics As Long) As Long
' Delete GDIP graphical area :
Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As Long) As Long
' Copy GDIP-Image to graphical area:
Private Declare Function GdipDrawImageRect Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
' Clear allocated bitmap memory from GDIP :
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
' Retrieve windows bitmap handle from GDIP-Image:
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
' Retrieve Windows-Icon-Handle from GDIP-Image:
Public Declare Function GdipCreateHICONFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
' Scaling GDIP-Image size:
Private Declare Function GdipGetImageThumbnail Lib "GDIPlus" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
' Retrieve GDIP-Image from Windows-Bitmap-Handle:
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, BITMAP As Long) As Long
' Retrieve GDIP-Image from Windows-Icon-Handle:
Private Declare Function GdipCreateBitmapFromHICON Lib "GDIPlus" (ByVal hicon As Long, BITMAP As Long) As Long
' Retrieve width of a GDIP-Image (Pixel):
Private Declare Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As Long, Width As Long) As Long
' Retrieve height of a GDIP-Image (Pixel):
Private Declare Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As Long, Height As Long) As Long
' Save GDIP-Image to file in seletable format:
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
' Save GDIP-Image in OLE-Stream with seletable format:
Private Declare Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As Long, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
' Retrieve GDIP-Image from OLE-Stream-Object:
Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal stream As IUnknown, Image As Long) As Long
' Create a gdip image from scratch
Private Declare Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
' Get the DC of an gdip image
Private Declare Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As Long, graphics As Long) As Long
' Blit the contents of an gdip image to another image DC using positioning
Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long

'-----------------------------------------------------------------------------------------
' Global module variable:
Private lGDIP As Long
Private bSharedLoad As Boolean
'-----------------------------------------------------------------------------------------


' Initialize GDI+
Private Function InitGDIP() As Boolean

    Dim TGDP As GDIPStartupInput
    Dim hMod As Long

    On Error GoTo PROC_ERR

    If lGDIP = 0 Then
        If IsNull(TempVars("GDIPlusHandle")) Then   ' If lGDIP is broken due to unhandled errors restore it from the Tempvars collection
            TGDP.GdiplusVersion = 1
            'MsgBox "Val(Application.Version)=" & Val(Application.Version)
            If Val(Application.Version) <> "15" Then
                MsgBox "This demo is for Access 2013 only!", vbCritical, "GDIPlusDemo2013"
                Stop
            End If
            hMod = GetModuleHandle("gdiplus.dll")   ' gdiplus.dll not yet loaded?
'            If hMod = 0 Then
'                If Val(Application.Version) = 14 Then   ' Distinguish between Office 12 (2007) and Office 14 (2010)
'                    hMod = LoadLibrary(Environ$("CommonProgramFiles") & "\Microsoft Shared\Office14\ogl.dll")
'                Else
'                    hMod = LoadLibrary(Environ$("CommonProgramFiles") & "\Microsoft Shared\Office12\ogl.dll")
'                End If
'                bSharedLoad = False
'            Else
                bSharedLoad = True
'            End If
            GdiplusStartup lGDIP, TGDP  ' Get a personal instance of gdiplus
            TempVars("GDIPlusHandle") = lGDIP
        Else
            lGDIP = TempVars("GDIPlusHandle")
        End If
        AutoShutDown
    End If
    InitGDIP = (lGDIP > 0)

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = 53 Then    ' File not found
        MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "InitGDIP"
        Debug.Print "Erl=" & Erl & " Err=" & Err & " " & Err.Description & vbCrLf & " in procedure InitGDIP"
        Stop
    Else
        MsgBox "Erl=" & Erl & " Err=" & Err & " " & Err.Description, vbCritical, "InitGDIP"
    End If
    Stop
    Resume Next

End Function

' Clear GDI+
Public Sub ShutDownGDIP()
    On Error GoTo 0
    If lGDIP <> 0 Then
        If KillTimer(0&, CLng(TempVars("TimerHandle"))) Then TempVars("TimerHandle") = 0
        GdiplusShutdown lGDIP
        lGDIP = 0
        TempVars("GDIPlusHandle") = Null
        If Not bSharedLoad Then FreeLibrary GetModuleHandle("gdiplus.dll")
    End If
End Sub

' Scheduled ShutDown of GDI+ handle to avoid memory leaks
Private Sub AutoShutDown()
    On Error GoTo 0
    ' Set to 5 seconds for next shutdown
    ' That's IMO appropriate for looped routines  - but configure for your own purposes
    If lGDIP <> 0 Then
        TempVars("TimerHandle") = SetTimer(0&, 0&, 5000, AddressOf TimerProc)
    End If
End Sub

' Callback for AutoShutDown
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error GoTo 0
    Debug.Print "GDI+ AutoShutDown", idEvent
    If TempVars("TimerHandle") <> 0 Then
        If KillTimer(0&, CLng(TempVars("TimerHandle"))) Then TempVars("TimerHandle") = 0
    End If
    ShutDownGDIP
End Sub

' Load image file with GDIP
' It's equivalent to the method LoadPicture() in OLE-Automation library (stdole2.tlb)
' Allowed format: bmp, gif, jp(e)g, tif, png, wmf, emf, ico
Public Function LoadPictureGDIP(sFileName As String) As StdPicture

    Dim hBmp As Long
    Dim hPic As Long

    On Error GoTo 0

    If Not InitGDIP Then Exit Function
    If GdipCreateBitmapFromFile(StrPtr(sFileName), hPic) = 0 Then
        GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&
        If hBmp <> 0 Then
            Set LoadPictureGDIP = BitmapToPicture(hBmp)
            GdipDisposeImage hPic
        End If
    End If

End Function

' Scale picture with GDIP
' A Picture object is commited, also the return value
' Width and Height of generatrix pictures in Width, Height
' bSharpen: TRUE=Thumb is additional sharpened
Public Function ResampleGDIP(ByVal Image As StdPicture, ByVal Width As Long, ByVal Height As Long, _
                      Optional bSharpen As Boolean = True) As StdPicture
    Dim lRes As Long
    Dim lBitmap As Long

    On Error GoTo 0

    If Not InitGDIP Then Exit Function
    
    If Image.type = 1 Then
        lRes = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap)
    Else
        lRes = GdipCreateBitmapFromHICON(Image.Handle, lBitmap)
    End If
    If lRes = 0 Then
        Dim lThumb As Long
        Dim hBitmap As Long

        lRes = GdipGetImageThumbnail(lBitmap, Width, Height, lThumb, 0, 0)
        If lRes = 0 Then
            If Image.type = 3 Then  ' Image-Type 3 is named : Icon!
                ' Convert with these GDI+ method :
                lRes = GdipCreateHICONFromBitmap(lThumb, hBitmap)
                Set ResampleGDIP = BitmapToPicture(hBitmap, True)
            Else
                lRes = GdipCreateHBITMAPFromBitmap(lThumb, hBitmap, 0)
                Set ResampleGDIP = BitmapToPicture(hBitmap)
            End If
            
            GdipDisposeImage lThumb
        End If
        GdipDisposeImage lBitmap
    End If

End Function

' Extract a part of an image
' x,y:           Left top corner of area to extract (pixel)
' Width, Height: Width and height of area to extract
' Return:        Image partly extracted
Private Function CropImage(ByVal Image As StdPicture, _
                   X As Long, Y As Long, _
                   Width As Long, Height As Long) As StdPicture
    Dim ret As Long
    Dim lBitmap As Long
    Dim lBitmap2 As Long
    Dim lGraph As Long
    Dim hBitmap As Long
    Dim sx As Long, sy As Long
    
    Const PixelFormat32bppARGB = &H26200A
    Const UnitPixel = 2

    On Error GoTo 0

    If Not InitGDIP Then Exit Function
    
    ret = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap)
    If ret = 0 Then
        ret = GdipGetImageWidth(lBitmap, sx)
        ret = GdipGetImageHeight(lBitmap, sy)
        If (X + Width) > sx Then Width = sx - X
        If (Y + Height) > sy Then Height = sy - Y
        ret = GdipCreateBitmapFromScan0(CLng(Width), CLng(Height), _
                    0, PixelFormat32bppARGB, ByVal 0&, lBitmap2)
        ret = GdipGetImageGraphicsContext(lBitmap2, lGraph)
        ret = GdipDrawImageRectRectI(lGraph, lBitmap, 0&, 0&, _
                    Width, Height, X, Y, Width, Height, UnitPixel)
        ret = GdipCreateHBITMAPFromBitmap(lBitmap2, hBitmap, 0)
        Set CropImage = BitmapToPicture(hBitmap)
        
        GdipDisposeImage lBitmap
        GdipDisposeImage lBitmap2
        GdipDeleteGraphics lGraph
    End If

End Function

' Retrieve Width and Height of a pictures in Pixel with GDIP
' Return value as user/defined type TSize (X/Y als Long)
Public Function GetDimensionsGDIP(ByVal Image As StdPicture) As TSize

    Dim lRes As Long
    Dim lBitmap As Long
    Dim X As Long, Y As Long

    On Error GoTo 0

    If Not InitGDIP Then Exit Function
    If Image Is Nothing Then Exit Function
    lRes = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap)
    If lRes = 0 Then
        GdipGetImageHeight lBitmap, Y
        GdipGetImageWidth lBitmap, X
        GetDimensionsGDIP.X = CDbl(X)
        GetDimensionsGDIP.Y = CDbl(Y)
        GdipDisposeImage lBitmap
    End If

End Function

' Save a bitmap as file (with format conversion!)
' image = StdPicture object
' sFile = complete file path
' PicType = pictypeBMP, pictypeGIF, pictypePNG oder pictypeJPG
' Quality: 0...100; (works only with pictypeJPG!)
' Returns TRUE if successful
Private Function SavePicGDIPlus(ByVal Image As StdPicture, sFile As String, _
                        PicType As PicFileType, Optional Quality As Long = 80) As Boolean

    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String

    On Error GoTo 0

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap) = 0 Then
        Select Case PicType
            Case pictypeBMP
                sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypeGIF
                sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypePNG
                sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypeJPG
                sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder
        If PicType = pictypeJPG Then
            TParams.count = 1
            With TParams.Parameter    ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(CLng(Quality))
            End With
        Else
            ' Different numbers of parameter between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
            If (PicType = pictypeGIF) Then TParams.count = 1 Else TParams.count = 0
        End If
        ' Save GDIP-Image to file :
        ret = GdipSaveImageToFile(lBitmap, StrPtr(sFile), TEncoder, TParams)
        GdipDisposeImage lBitmap
        DoEvents
        ' Function returns True, if generated file actually exists:
        SavePicGDIPlus = (Dir(sFile) <> "")
    End If

End Function

' This procedure is similar to the above (see Parameter), the different is,
' that nothing is stored as a file, but a conversion is executed
' using a OLE-Stream-Object to an Byte-Array .
Public Function ArrayFromPicture(ByVal Image As Object, PicType As PicFileType, Optional Quality As Long = 80) As Byte()

    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String
    Dim IStm As IUnknown

    On Error GoTo 0

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap) = 0 Then
        Select Case PicType    ' Choose GDIP-Format-Encoders CLSID:
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder

        If PicType = pictypeJPG Then    ' If JPG, then set additional parameter
                                        ' to apply quality level
            TParams.count = 1
            With TParams.Parameter    ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(CLng(Quality))
            End With
        Else
            ' Different number of parameters between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
            If (PicType = pictypeGIF) Then TParams.count = 1 Else TParams.count = 0
        End If

        ret = CreateStreamOnHGlobal(0&, 1, IStm)    ' Create stream
        ' Save GDIP-Image to stream :
        ret = GdipSaveImageToStream(lBitmap, IStm, TEncoder, TParams)
        If ret = 0 Then
            Dim hMem As Long, LSize As Long, lpMem As Long
            Dim abData() As Byte

            ret = GetHGlobalFromStream(IStm, hMem)    ' Get memory handle from stream
            If ret = 0 Then
                LSize = GlobalSize(hMem)
                lpMem = GlobalLock(hMem)   ' Get access to memory
                ReDim abData(LSize - 1)    ' Arrays dimension
                ' Commit memory stack from streams :
                CopyMemory abData(0), ByVal lpMem, LSize
                GlobalUnlock hMem   ' Lock memory
                ArrayFromPicture = abData   ' Result
            End If

            Set IStm = Nothing  ' Clean
        End If

        GdipDisposeImage lBitmap    ' Clear GDIP-Image-Memory
    End If

End Function

' Create a picture object from an OLE Field (BLOB, long binary)
' strTable:              Table containing OLE field with picture contents
' strNameField:          Name field to identify record
' strName:               Unique name of the picture in Name field
' strOLEField:           Name of OLE field in table
' ? OLEFieldToPicture("tblOLE","ImageName","cloudy","Blob").Width
Public Function OLEFieldToPicture(strTable As String, _
                                  strNameField As String, _
                                  strName As String, _
                                  strOLEField As String) As StdPicture
    On Error GoTo 0

    Dim rst As Recordset2
    Set rst = CurrentDb.OpenRecordset("SELECT " & strOLEField & " FROM " & strTable & " WHERE " & strNameField & "='" & strName & "'", dbOpenDynaset)

    If Not rst.EOF Then
        Set OLEFieldToPicture = ArrayToPicture(rst(strOLEField).Value)
    End If
    rst.Close
    Set rst = Nothing
    
End Function

' Create a picture object from an Access 2007 attachment
' strTable:              Table containing picture file attachments
' strAttachmentField:    Name of the attachment column in the table
' strImage:              Name of the image to search in the attachment records
' ? AttachmentToPicture("tblImages","Image","check16.png").Width
Public Function AttachmentToPicture(strTable As String, _
                                    strAttachmentField As String, _
                                    strImage As String) As StdPicture
    Dim strSQL As String
    Dim bin() As Byte
    Dim nOffset As Long
    Dim nSize As Long

    On Error GoTo 0

    strSQL = "SELECT " & strTable & "." & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & _
             " WHERE " & strTable & "." & strAttachmentField & ".FileName='" & strImage & "'"
    On Error Resume Next
    bin = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenSnapshot)(0)
    If Err.Number = 0 Then
        Dim bin2() As Byte
        nOffset = bin(0)    ' First byte of Field2.FileData identifies offset to the file data block
        nSize = UBound(bin)
        ReDim bin2(nSize - nOffset)
        CopyMemory bin2(0), bin(nOffset), nSize - nOffset   ' Copy file into new byte array starting at nOffset
        Set AttachmentToPicture = ArrayToPicture(bin2)
        Erase bin2
        Erase bin
    End If
End Function

' Create an OLE-Picture from Byte-Array PicBin()
Public Function ArrayToPicture(ByRef PicBin() As Byte) As StdPicture

    Dim IStm As IUnknown
    Dim lBitmap As Long
    Dim hBmp As Long
    Dim ret As Long

    On Error GoTo 0

    If Not InitGDIP Then Exit Function

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)  ' Create stream from memory stack
    If ret = 0 Then    ' OK, start GDIP :
        ' Convert stream to GDIP-Image :
        ret = GdipLoadImageFromStream(IStm, lBitmap)
        If ret = 0 Then
            ' Get Windows-Bitmap from GDIP-Image:
            GdipCreateHBITMAPFromBitmap lBitmap, hBmp, 0&
            If hBmp <> 0 Then
                ' Convert bitmap to picture object :
                Set ArrayToPicture = BitmapToPicture(hBmp)
            End If
        End If
        ' Clear memory ...
        GdipDisposeImage lBitmap
    End If

End Function

' Help function to get a OLE-Picture from Windows-Bitmap-Handle
' If bIsIcon = TRUE, an Icon-Handle is commited
Private Function BitmapToPicture(ByVal hBmp As Long, Optional bIsIcon As Boolean = False) As StdPicture

    On Error GoTo 0

    Dim TPicConv As PICTDESC, UID As GUID

    With TPicConv
        If bIsIcon Then
            .cbSizeOfStruct = 16
            .PicType = 3    ' PicType Icon
        Else
            .cbSizeOfStruct = Len(TPicConv)
            .PicType = 1    ' PicType Bitmap
        End If
        .hImage = hBmp
    End With

    CLSIDFromString StrPtr(GUID_IPicture), UID
    OleCreatePictureIndirect TPicConv, UID, True, BitmapToPicture

End Function