Option Compare Database
Option Explicit

'-------------------------------------------------
'    Picture functions using GDIPlus-API (GDIP)   |
'   (c) mossSOFT / Sascha Trowitzsch rev. 06/2013 |
'             http://www.mosstools.de             |
'-------------------------------------------------|
'          *  Office 2013+ GDayClass  *           |
'            (c) 2014 Peter F. Ennis              |
'-------------------------------------------------|


' Reference to library "OLE Automation" (stdole) needed!

Public Const GUID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"    'IPicture

'User-defined types: ----------------------------------------------------------------------

Public Enum PicFileType
    pictypeBMP = 1
    pictypeGIF = 2
    pictypePNG = 3
    pictypeJPG = 4
End Enum

Private Enum GpUnit
   UnitWorld = 0        ' World coordinate (non-physical unit)
   UnitDisplay = 1      ' Variable -- for PageTransform only
   UnitPixel = 2        ' Each unit is one device pixel.
   UnitPoint = 3        ' Each unit is a printer's point, or 1/72 inch.
   UnitInch = 4         ' Each unit is 1 inch.
   UnitDocument = 5     ' Each unit is 1/300 inch.
   UnitMillimeter = 6   ' Each unit is 1 millimeter.
End Enum

Private Enum PixelFormat
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    pixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
    PixelFormatMax = 15 '&HF
End Enum

Private Enum ImageLockMode
   ImageLockModeRead = &H1
   ImageLockModewrite = &H2
   ImageLockModeUserInputBuf = &H4
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type TSize
    X As Double
    Y As Double
End Type

Private Type RECT
    Bottom As Long
    Left As Long
    Right As Long
    Top As Long
End Type

Private Type RECTL
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

Private Type BitmapData
   Width As Long
   Height As Long
   stride As Long
   PixelFormat As Long
   scan0 As Long
   Reserved As Long
End Type

'Common API-Declarations: ----------------------------------------------------------------------------

'Convert a windows bitmap to OLE-Picture :
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
'Retrieve GUID-Type from string :
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pCLSID As GUID) As Long

'Memory functions:
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Byte, ByVal Length As Long)

'Modules API:
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'Timer API:
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


'OLE-Stream functions :
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As Long) As Long

'GDIPlus Flat-API declarations ----------------------------------------------------------------------------

'*Remark:
'          We use a special gdi+ version here that comes with Office 2007/2010! (program files\common files\microsoft shared\office1x\ogl.dll)
'          Benefit: No need to load a separate dll because ogl.dll is normally already loaded by Office 2007/2010.
'          ogl.dll is identical to the gdiplus.dll (V1.1) used in Vista
'Remark 2: This DLL is only installed by Office Setup (and also Access Runtime) if OS = WinXP.
'          On Vista or Win7 Office 2010 uses the built-in GDIPLUS.DLL !


'OGL.DLL library declarations:

'Initialization OGL:
Private Declare Function GdiplusStartup_O Lib "ogl" Alias "GdiplusStartup" (token As Long, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
'Tear down GDIP:
Private Declare Function GdiplusShutdown_O Lib "ogl" Alias "GdiplusShutdown" (ByVal token As Long) As Long
'Load GDIP-Image from file :
Private Declare Function GdipCreateBitmapFromFile_O Lib "ogl" Alias "GdipCreateBitmapFromFile" (ByVal FileName As Long, BITMAP As Long) As Long
'Create GDIP- graphical area from Windows-DeviceContext:
Private Declare Function GdipCreateFromHDC_O Lib "ogl" Alias "GdipCreateFromHDC" (ByVal hdc As Long, GpGraphics As Long) As Long
'Delete GDIP graphical area :
Private Declare Function GdipDeleteGraphics_O Lib "ogl" Alias "GdipDeleteGraphics" (ByVal graphics As Long) As Long
'Copy GDIP-Image to graphical area:
Private Declare Function GdipDrawImageRect_O Lib "ogl" Alias "GdipDrawImageRect" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
'Clear allocated bitmap memory from GDIP :
Private Declare Function GdipDisposeImage_O Lib "ogl" Alias "GdipDisposeImage" (ByVal Image As Long) As Long
'Retrieve windows bitmap handle from GDIP-Image:
Private Declare Function GdipCreateHBITMAPFromBitmap_O Lib "ogl" Alias "GdipCreateHBITMAPFromBitmap" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
'Retrieve Windows-Icon-Handle from GDIP-Image:
Public Declare Function GdipCreateHICONFromBitmap_O Lib "ogl" Alias "GdipCreateHICONFromBitmap" (ByVal BITMAP As Long, hbmReturn As Long) As Long
'Scaling GDIP-Image size:
Private Declare Function GdipGetImageThumbnail_O Lib "ogl" Alias "GdipGetImageThumbnail" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
'Retrieve GDIP-Image from Windows-Bitmap-Handle:
Private Declare Function GdipCreateBitmapFromHBITMAP_O Lib "ogl" Alias "GdipCreateBitmapFromHBITMAP" (ByVal hbm As Long, ByVal hpal As Long, BITMAP As Long) As Long
'Retrieve GDIP-Image from Windows-Icon-Handle:
Private Declare Function GdipCreateBitmapFromHICON_O Lib "ogl" Alias "GdipCreateBitmapFromHICON" (ByVal hicon As Long, BITMAP As Long) As Long
'Retrieve width of a GDIP-Image (Pixel):
Private Declare Function GdipGetImageWidth_O Lib "ogl" Alias "GdipGetImageWidth" (ByVal Image As Long, Width As Long) As Long
'Retrieve height of a GDIP-Image (Pixel):
Private Declare Function GdipGetImageHeight_O Lib "ogl" Alias "GdipGetImageHeight" (ByVal Image As Long, Height As Long) As Long
'Save GDIP-Image to file in seletable format:
Private Declare Function GdipSaveImageToFile_O Lib "ogl" Alias "GdipSaveImageToFile" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
'Save GDIP-Image in OLE-Stream with seletable format:
Private Declare Function GdipSaveImageToStream_O Lib "ogl" Alias "GdipSaveImageToStream" (ByVal Image As Long, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
'Retrieve GDIP-Image from OLE-Stream-Object:
Private Declare Function GdipLoadImageFromStream_O Lib "ogl" Alias "GdipLoadImageFromStream" (ByVal stream As IUnknown, Image As Long) As Long
'Create a gdip image from scratch
Private Declare Function GdipCreateBitmapFromScan0_O Lib "ogl" Alias "GdipCreateBitmapFromScan0" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
'Get the DC of an gdip image
Private Declare Function GdipGetImageGraphicsContext_O Lib "ogl" Alias "GdipGetImageGraphicsContext" (ByVal Image As Long, graphics As Long) As Long
'Blit the contents of an gdip image to another image DC using positioning
Private Declare Function GdipDrawImageRectRectI_O Lib "ogl" Alias "GdipDrawImageRectRectI" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
'Duplicates a gdiplus image object
Private Declare Function GdipCloneImage_O Lib "ogl" Alias "GdipCloneImage" (ByVal Image As Long, cloneImage As Long) As Long
'Clear device context and set background color
Private Declare Function GdipGraphicsClear_O Lib "ogl" Alias "GdipGraphicsClear" (ByVal graphics As Long, ByVal LColor As Long) As Long
'Suspend image to work with its data (pixels)
Private Declare Function GdipBitmapLockBits_O Lib "ogl" Alias "GdipBitmapLockBits" (ByVal BITMAP As Long, RECT As RECTL, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
'Continue to use altered image in GDIP
Private Declare Function GdipBitmapUnlockBits_O Lib "ogl" Alias "GdipBitmapUnlockBits" (ByVal BITMAP As Long, lockedBitmapData As BitmapData) As Long


'Same for simple GDIPLUS.DLL:
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, BITMAP As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, GpGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
Private Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, BITMAP As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hicon As Long, BITMAP As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As IUnknown, Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, graphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As Long, ByVal LColor As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, RECT As RECTL, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, lockedBitmapData As BitmapData) As Long


'-----------------------------------------------------------------------------------------
'Global module variables:
Private lGDIP As Long           'GDIPLus object instance
Private bSharedLoad As Boolean  'Is gdiplus.dll or ogl.dll already loaded by Access? (In this case do not FreeLibrary module)
Private bUseOGL As Boolean      'If True use ogl.dll, otherwise gdiplus.dll
Private IsGDI11 As Boolean      'Is GDIPLUS version 1.1 or 1.0? (1.1 supports effects like Sharpen etc.)
Private lTimer As Long          'Timer Handle for AutoShutdown
'Be sure to have error handlers in all your VBA procedures since unhandled errors clear the above variables
'This may cause instableties or even crashes due to memory leaks in gdiplus!
'-----------------------------------------------------------------------------------------

Function GetGDIPVersion() As Boolean
    Dim hMod As Long
    Select Case Application.Version
    Case "11.0", "15.0" 'A2003, A2013
        bUseOGL = False
        hMod = GetModuleHandle("gdiplus.dll")
        If hMod = 0 Then
            hMod = LoadLibrary("gdiplus.dll")
        Else
            bSharedLoad = True
        End If
        Dim lAddr As Long
        lAddr = GetProcAddress(hMod, "GdipCreateEffect")    'Check if effect section is supported by GDIPLUS module (=V 1.1)
        IsGDI11 = (lAddr <> 0)
        
    Case "12.0" 'A2007
        bUseOGL = True
        IsGDI11 = True
        hMod = GetModuleHandle("ogl.dll")
        If hMod = 0 Then
            hMod = LoadLibrary(Environ$("CommonProgramFiles") & "\Microsoft Shared\Office12\ogl.dll")
        Else
            bSharedLoad = True
        End If
        
    Case "14.0" 'A2010
        IsGDI11 = True
        'Office 2010 Setup only installs the OGL module, if OS <> Vista or Win7!
        'Check here for existance:
        hMod = GetModuleHandle("ogl.dll")   'Attempt Shared OGL
        If hMod <> 0 Then
            bUseOGL = True
            bSharedLoad = True
        Else
            hMod = GetModuleHandle("gdiplus.dll")   'Attempt Shared GDIPLUS
            If hMod <> 0 Then bSharedLoad = True
        End If
        If hMod = 0 Then    'Not Shared, so load the library...
            hMod = LoadLibrary(Environ$("CommonProgramFiles") & "\Microsoft Shared\Office14\ogl.dll")
            If hMod <> 0 Then
                bUseOGL = True
            Else
                hMod = LoadLibrary("gdiplus.dll")   'OGL does not exist, so load Vistas or Win7s gdiplus.dll (= always V 1.1)
            End If
        End If
    End Select
    GetGDIPVersion = (hMod <> 0)    'Valid only if we could receive any module handle
End Function

'Initialize GDI+
Function InitGDIP() As Boolean
    Dim TGDP As GDIPStartupInput
    Dim hMod As Long

    If lGDIP = 0 Then
        If GetGDIPVersion Then  'Distinguish between Office and OS versions
            TGDP.GdiplusVersion = 1
            If bUseOGL Then 'Get a personal instance of gdiplus:
                GdiplusStartup_O lGDIP, TGDP
            Else
                GdiplusStartup lGDIP, TGDP
            End If
            If lGDIP <> 0 Then AutoShutDown
        End If
    End If
    InitGDIP = (lGDIP > 0)
End Function

'Clear GDI+
Sub ShutDownGDIP()
    If lGDIP <> 0 Then
        If KillTimer(0&, lTimer) Then lTimer = 0
        If bUseOGL Then GdiplusShutdown_O lGDIP Else GdiplusShutdown lGDIP
        lGDIP = 0
        If Not bSharedLoad Then
            If bUseOGL Then FreeLibrary GetModuleHandle("ogl.dll") Else FreeLibrary GetModuleHandle("gdiplus.dll")
        End If
    End If
End Sub

'Scheduled ShutDown of GDI+ handle to avoid memory leaks
Private Sub AutoShutDown()
    'Set to 5 seconds for next shutdown
    'That's IMO appropriate for looped routines  - but configure for your own purposes
    If lGDIP <> 0 Then
        lTimer = SetTimer(0&, 0&, 5000, AddressOf TimerProc)
    End If
End Sub

'Callback for AutoShutDown
Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Debug.Print "GDI+ AutoShutDown", idEvent
    If lTimer <> 0 Then
        If KillTimer(0&, lTimer) Then lTimer = 0
    End If
    ShutDownGDIP
End Sub

Function UsesOGL() As Boolean
    If Not InitGDIP Then Exit Function
    UsesOGL = bUseOGL
End Function

'Load image file with GDIP
'It's equivalent to the method LoadPicture() in OLE-Automation library (stdole2.tlb)
'Allowed format: bmp, gif, jp(e)g, tif, png, wmf, emf, ico
Function LoadPictureGDIP(sFileName As String) As StdPicture
    Dim hBmp As Long
    Dim hPic As Long

    If Not InitGDIP Then Exit Function
    If bUseOGL Then
        Set LoadPictureGDIP = LoadPictureGDIP_O(sFileName)
    Else
        If GdipCreateBitmapFromFile(StrPtr(sFileName), hPic) = 0 Then
            GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&
            If hBmp <> 0 Then
                Set LoadPictureGDIP = BitmapToPicture(hBmp)
                GdipDisposeImage hPic
            End If
        End If
    End If

End Function

'Scale picture with GDIP
'A Picture object is commited, also the return value
'Width and Height of generatrix pictures in Width, Height
Public Function ResampleGDIP(ByVal Image As StdPicture, ByVal Width As Long, ByVal Height As Long) As StdPicture
    Dim lRes As Long
    Dim lBitmap As Long

    If Not InitGDIP Then Exit Function

    If bUseOGL Then
        Set ResampleGDIP = ResampleGDIP_O(Image, Width, Height)
    Else
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
                If Image.type = 3 Then  'Image-Type 3 is named : Icon!
                    'Convert with these GDI+ method :
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
    End If

End Function

'Extract a part of an image
'x,y:   Left top corner of area to extract (pixel)
'Width, Height: Width and height of area to extract
'Return:    Image partly extracted
Function CropImage(ByVal Image As StdPicture, _
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

    If Not InitGDIP Then Exit Function

    If bUseOGL Then
        Set CropImage = CropImage_O(Image, X, Y, Width, Height)
    Else
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
    End If

End Function

'Retrieve Width and Height of a pictures in Pixel with GDIP
'Return value as user/defined type TSize (X/Y als Long)
Function GetDimensionsGDIP(ByVal Image As StdPicture) As TSize
    Dim lRes As Long
    Dim lBitmap As Long
    Dim X As Long, Y As Long

    If Not InitGDIP Then Exit Function
    If Image Is Nothing Then Exit Function
    If bUseOGL Then
        GetDimensionsGDIP = GetDimensionsGDIP_O(Image)
    Else
        lRes = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap)
        If lRes = 0 Then
            GdipGetImageHeight lBitmap, Y
            GdipGetImageWidth lBitmap, X
            GetDimensionsGDIP.X = CDbl(X)
            GetDimensionsGDIP.Y = CDbl(Y)
            GdipDisposeImage lBitmap
        End If
    End If

End Function

'Save a bitmap as file (with format conversion!)
'image = StdPicture object
'sFile = complete file path
'PicType = pictypeBMP, pictypeGIF, pictypePNG oder pictypeJPG
'Quality: 0...100; (works only with pictypeJPG!)
'Returns TRUE if successful
Function SavePicGDIPlus(ByVal Image As StdPicture, sFile As String, _
                        PicType As PicFileType, Optional Quality As Long = 80) As Boolean
    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String

    If Not InitGDIP Then Exit Function

    If bUseOGL Then
        SavePicGDIPlus = SavePicGDIPlus_O(Image, sFile, pictypeBMP, Quality)
    Else
        If GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap) = 0 Then
            Select Case PicType
            Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
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
                'Different numbers of parameter between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
                If (PicType = pictypeGIF) Then TParams.count = 1 Else TParams.count = 0
                If (PicType = pictypePNG) And IsGDI11 Then TParams.count = 1
            End If
            'Save GDIP-Image to file :
            ret = GdipSaveImageToFile(lBitmap, StrPtr(sFile), TEncoder, TParams)
            GdipDisposeImage lBitmap
            DoEvents
            'Function returns True, if generated file actually exists:
            SavePicGDIPlus = (Dir(sFile) <> "")
        End If
    End If

End Function

'This procedure is similar to the above (see Parameter), the different is,
'that nothing is stored as a file, but a conversion is executed
'using a OLE-Stream-Object to an Byte-Array .
Function ArrayFromPicture(ByVal Image As Object, PicType As PicFileType, Optional Quality As Long = 80) As Byte()
    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String
    Dim IStm As IUnknown

    If Not InitGDIP Then Exit Function

    If bUseOGL Then
        ArrayFromPicture = ArrayFromPicture_O(Image, pictypeBMP, Quality)
    Else
        If GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap) = 0 Then
            Select Case PicType    'Choose GDIP-Format-Encoders CLSID:
            Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
            Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
            End Select
            CLSIDFromString StrPtr(sType), TEncoder

            If PicType = pictypeJPG Then    'If JPG, then set additional parameter
                ' to apply quality level
                TParams.count = 1
                With TParams.Parameter    ' Quality
                    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                    .NumberOfValues = 1
                    .type = 4
                    .Value = VarPtr(CLng(Quality))
                End With
            Else
                'Different number of parameters between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
                If (PicType = pictypeGIF) Then TParams.count = 1 Else TParams.count = 0
                'For PNGs and in case of GDIPlus1.1 there are two parameters required (bugfix) :
                If (PicType = pictypePNG) And IsGDI11 Then TParams.count = 1
            End If

            ret = CreateStreamOnHGlobal(0&, 1, IStm)    'Create stream
            'Save GDIP-Image to stream :
            ret = GdipSaveImageToStream(lBitmap, IStm, TEncoder, TParams)
            If ret = 0 Then
                Dim hMem As Long, LSize As Long, lpMem As Long
                Dim abData() As Byte

                ret = GetHGlobalFromStream(IStm, hMem)    'Get memory handle from stream
                If ret = 0 Then
                    LSize = GlobalSize(hMem)
                    lpMem = GlobalLock(hMem)   'Get access to memory
                    ReDim abData(LSize - 1)    'Arrays dimension
                    'Commit memory stack from streams :
                    CopyMemory abData(0), ByVal lpMem, LSize
                    GlobalUnlock hMem   'Lock memory
                    ArrayFromPicture = abData   'Result
                End If

                Set IStm = Nothing  'Clean
            End If

            GdipDisposeImage lBitmap    'Clear GDIP-Image-Memory
        End If
    End If
End Function

'Create a picture object from an OLE Field (BLOB, long binary)
'strTable:              Table containing OLE field with picture contents
'strNameField:          Name field to identify record
'strName:               Unique name of the picture in Name field
'strOLEField:           Name of OLE field in table
'? OLEFieldToPicture("tblOLE","ImageName","cloudy","Blob").Width
Public Function OLEFieldToPicture(ByVal strTable As String, _
                                  ByVal strNameField As String, _
                                  ByVal strName As String, _
                                  ByVal strOLEField As String) As StdPicture
    On Error GoTo 0

    Dim rst As Recordset2
    Set rst = CurrentDb.OpenRecordset("SELECT " & strOLEField & " FROM " & strTable & " WHERE " & strNameField & "='" & strName & "'", dbOpenDynaset)

    If Not rst.EOF Then
        Set OLEFieldToPicture = ArrayToPicture(rst(strOLEField).Value)
    End If
    rst.Close
    Set rst = Nothing
    
End Function

'Create a picture object from an Access 2007 attachment
'strTable:              Table containing picture file attachments
'strAttachmentField:    Name of the attachment column in the table
'strImage:              Name of the image to search in the attachment records
'? AttachmentToPicture("ribbonimages","imageblob","cloudy.png").Width
Public Function AttachmentToPicture(strTable As String, strAttachmentField As String, strImage As String) As StdPicture
    Dim strSQL As String
    Dim bin() As Byte
    Dim nOffset As Long
    Dim nSize As Long

    strSQL = "SELECT " & strTable & "." & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & _
             " WHERE " & strTable & "." & strAttachmentField & ".FileName='" & strImage & "'"
    On Error Resume Next
    bin = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenSnapshot)(0)
    If Err.Number = 0 Then
        Dim bin2() As Byte
        nOffset = bin(0)    'First byte of Field2.FileData identifies offset to the file data block
        nSize = UBound(bin)
        ReDim bin2(nSize - nOffset)
        CopyMemory bin2(0), bin(nOffset), nSize - nOffset   'Copy file into new byte array starting at nOffset
        If bUseOGL Then
            Set AttachmentToPicture = ArrayToPicture_O(bin2)
        Else
            Set AttachmentToPicture = ArrayToPicture(bin2)
        End If
        Erase bin2
        Erase bin
    End If
End Function

Public Function PicFromField(ByVal picField As DAO.Field, Optional FlattenColor As Variant = &HFFFFFFFF) As StdPicture
    Dim arrBin() As Byte
    Dim LSize As Long

    On Error GoTo Fehler

    LSize = picField.FieldSize
    If LSize > 0 Then
        arrBin() = picField.GetChunk(0, LSize)
        Set PicFromField = ArrayToPicture(arrBin, FlattenColor)
    End If

Ende:
    Erase arrBin
    Exit Function

Fehler:
    MsgBox Err.Description, vbCritical
    Resume Ende
End Function

'Create an OLE-Picture from Byte-Array PicBin()
Public Function ArrayToPicture(ByRef PicBin() As Byte, Optional FlattenColor As Variant) As StdPicture
    Dim IStm As IUnknown
    Dim lBitmap As Long
    Dim hBmp As Long
    Dim ret As Long

    If Not InitGDIP Then Exit Function

    If bUseOGL Then
        Set ArrayToPicture = ArrayToPicture_O(PicBin, FlattenColor)
    Else
        ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)    'Create stream from memory stack
        If ret = 0 Then    'OK, start GDIP :
            'Convert stream to GDIP-Image :
            ret = GdipLoadImageFromStream(IStm, lBitmap)
            If ret = 0 Then
                If Not IsMissing(FlattenColor) Then
                    Dim lBitmap2 As Long
                    Dim lGraph As Long
                    Dim W As Long, H As Long

                    ret = GdipCloneImage(lBitmap, lBitmap2)
                    ret = GdipGetImageGraphicsContext(lBitmap2, lGraph)
                    If ret = 0 Then
                        ret = GdipGetImageWidth(lBitmap, W)
                        ret = GdipGetImageHeight(lBitmap, H)
                        ret = GdipGraphicsClear(lGraph, CLng(FlattenColor))
                        ret = GdipDrawImageRectRectI(lGraph, lBitmap, 0, 0, W, H, 0, 0, W, H, _
                         UnitPixel, 0, 0)
                    End If
                    GdipCreateHBITMAPFromBitmap lBitmap2, hBmp, 0&
                    GdipDeleteGraphics lGraph
                Else
                    GdipCreateHBITMAPFromBitmap lBitmap, hBmp, 0&
                End If
                If hBmp <> 0 Then
                    'Convert bitmap to picture object :
                    Set ArrayToPicture = BitmapToPicture(hBmp)
                End If
            End If
            'Clear memory ...
            GdipDisposeImage lBitmap
        End If
    End If

End Function

Function MaskFromPicture(ByVal Image As StdPicture, Optional TransColor As Variant) As StdPicture
    Dim lBitmap As Long
    Dim hBitmap As Long
    Dim W As Long, H As Long
    Dim bytes() As Long
    Dim BD As BitmapData
    Dim rct As RECTL
    Dim X As Long, Y As Long
    Dim AlphaColor As Long
    Dim ret As Long

    If Not InitGDIP Then Exit Function

    If bUseOGL Then
        Set MaskFromPicture = MaskFromPicture_O(Image, TransColor)
    Else
        ret = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap)
        If ret = 0 Then
            ret = GdipGetImageWidth(lBitmap, W)
            ret = GdipGetImageHeight(lBitmap, H)
            With rct
                .Left = 0
                .Top = H
                .Right = W
                .Bottom = 0
            End With
            ReDim bytes(W, H)
            With BD
                .Width = W
                .Height = H
                .PixelFormat = PixelFormat32bppARGB
                .stride = 4 * CLng(.Width + 1)
                .scan0 = VarPtr(bytes(0, 0))
            End With
            ret = GdipBitmapLockBits(lBitmap, rct, ImageLockModeRead Or _
                                                   ImageLockModeUserInputBuf Or ImageLockModewrite, PixelFormat32bppARGB, BD)
            If IsMissing(TransColor) Then
                AlphaColor = bytes(0, 0)
            Else
                AlphaColor = CLng(TransColor)
            End If
            For X = 0 To W
                For Y = 0 To H
                    If bytes(X, Y) = AlphaColor Then bytes(X, Y) = &HFFFFFF Else bytes(X, Y) = &H0
                Next Y
            Next X

            ret = GdipBitmapUnlockBits(lBitmap, BD)
            GdipCreateHBITMAPFromBitmap lBitmap, hBitmap, 0&
            Set MaskFromPicture = BitmapToPicture(hBitmap)
            GdipDisposeImage lBitmap
        End If
    End If

End Function

'Helper function to get a OLE-Picture from Windows-Bitmap-Handle
'If bIsIcon = TRUE, an Icon-Handle is commited
Function BitmapToPicture(ByVal hBmp As Long, Optional bIsIcon As Boolean = False) As StdPicture
    Dim TPicConv As PICTDESC, UID As GUID

    With TPicConv
        If bIsIcon Then
            .cbSizeOfStruct = 16
            .PicType = 3    'PicType Icon
        Else
            .cbSizeOfStruct = Len(TPicConv)
            .PicType = 1    'PicType Bitmap
        End If
        .hImage = hBmp
    End With

    CLSIDFromString StrPtr(GUID_IPicture), UID
    OleCreatePictureIndirect TPicConv, UID, True, BitmapToPicture

End Function


'--------------------------------------------------------------------------------------------------
'Following the same procedures using the OGL library
'The procedure names are the same but ending with "_O" here
'(for comments see procs above)

Function LoadPictureGDIP_O(sFileName As String) As StdPicture
    Dim hBmp As Long
    Dim hPic As Long

    If Not InitGDIP Then Exit Function
    If GdipCreateBitmapFromFile_O(StrPtr(sFileName), hPic) = 0 Then
        GdipCreateHBITMAPFromBitmap_O hPic, hBmp, 0&
        If hBmp <> 0 Then
            Set LoadPictureGDIP_O = BitmapToPicture(hBmp)
            GdipDisposeImage_O hPic
        End If
    End If

End Function

Function ResampleGDIP_O(ByVal Image As StdPicture, ByVal Width As Long, ByVal Height As Long) As StdPicture
    Dim lRes As Long
    Dim lBitmap As Long

    If Not InitGDIP Then Exit Function

    If Image.type = 1 Then
        lRes = GdipCreateBitmapFromHBITMAP_O(Image.Handle, 0, lBitmap)
    Else
        lRes = GdipCreateBitmapFromHICON_O(Image.Handle, lBitmap)
    End If
    If lRes = 0 Then
        Dim lThumb As Long
        Dim hBitmap As Long

        lRes = GdipGetImageThumbnail_O(lBitmap, Width, Height, lThumb, 0, 0)
        If lRes = 0 Then
            If Image.type = 3 Then
                lRes = GdipCreateHICONFromBitmap_O(lThumb, hBitmap)
                Set ResampleGDIP_O = BitmapToPicture(hBitmap, True)
            Else
                lRes = GdipCreateHBITMAPFromBitmap_O(lThumb, hBitmap, 0)
                Set ResampleGDIP_O = BitmapToPicture(hBitmap)
            End If

            GdipDisposeImage_O lThumb
        End If
        GdipDisposeImage_O lBitmap
    End If

End Function

Function CropImage_O(ByVal Image As StdPicture, _
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

    If Not InitGDIP Then Exit Function

    ret = GdipCreateBitmapFromHBITMAP_O(Image.Handle, 0, lBitmap)
    If ret = 0 Then
        ret = GdipGetImageWidth_O(lBitmap, sx)
        ret = GdipGetImageHeight_O(lBitmap, sy)
        If (X + Width) > sx Then Width = sx - X
        If (Y + Height) > sy Then Height = sy - Y
        ret = GdipCreateBitmapFromScan0_O(CLng(Width), CLng(Height), _
                    0, PixelFormat32bppARGB, ByVal 0&, lBitmap2)
        ret = GdipGetImageGraphicsContext_O(lBitmap2, lGraph)
        ret = GdipDrawImageRectRectI_O(lGraph, lBitmap, 0&, 0&, _
                    Width, Height, X, Y, Width, Height, UnitPixel)
        ret = GdipCreateHBITMAPFromBitmap_O(lBitmap2, hBitmap, 0)
        Set CropImage_O = BitmapToPicture(hBitmap)

        GdipDisposeImage_O lBitmap
        GdipDisposeImage_O lBitmap2
        GdipDeleteGraphics_O lGraph
    End If

End Function

Function GetDimensionsGDIP_O(ByVal Image As StdPicture) As TSize
    Dim lRes As Long
    Dim lBitmap As Long
    Dim X As Long, Y As Long

    If Not InitGDIP Then Exit Function
    If Image Is Nothing Then Exit Function
    lRes = GdipCreateBitmapFromHBITMAP_O(Image.Handle, 0, lBitmap)
    If lRes = 0 Then
        GdipGetImageHeight_O lBitmap, Y
        GdipGetImageWidth_O lBitmap, X
        GetDimensionsGDIP_O.X = CDbl(X)
        GetDimensionsGDIP_O.Y = CDbl(Y)
        GdipDisposeImage_O lBitmap
    End If

End Function

Function SavePicGDIPlus_O(ByVal Image As StdPicture, sFile As String, _
                        PicType As PicFileType, Optional Quality As Long = 80) As Boolean
    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP_O(Image.Handle, 0, lBitmap) = 0 Then
        Select Case PicType
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder
        If PicType = pictypeJPG Then
            TParams.count = 1
            With TParams.Parameter
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(CLng(Quality))
            End With
        Else
            If (PicType = pictypeGIF) Then TParams.count = 1 Else TParams.count = 0
            If (PicType = pictypePNG) And IsGDI11 Then TParams.count = 1
        End If
        ret = GdipSaveImageToFile_O(lBitmap, StrPtr(sFile), TEncoder, TParams)
        GdipDisposeImage_O lBitmap
        DoEvents
        SavePicGDIPlus_O = (Dir(sFile) <> "")
    End If

End Function

Function ArrayFromPicture_O(ByVal Image As Object, PicType As PicFileType, Optional Quality As Long = 80) As Byte()
    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String
    Dim IStm As IUnknown

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP_O(Image.Handle, 0, lBitmap) = 0 Then
        Select Case PicType
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder

        If PicType = pictypeJPG Then
            TParams.count = 1
            With TParams.Parameter
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(CLng(Quality))
            End With
        Else
            If (PicType = pictypeGIF) Then TParams.count = 1 Else TParams.count = 0
            If (PicType = pictypePNG) And IsGDI11 Then TParams.count = 1
        End If

        ret = CreateStreamOnHGlobal(0&, 1, IStm)
        ret = GdipSaveImageToStream_O(lBitmap, IStm, TEncoder, TParams)
        If ret = 0 Then
            Dim hMem As Long, LSize As Long, lpMem As Long
            Dim abData() As Byte

            ret = GetHGlobalFromStream(IStm, hMem)
            If ret = 0 Then
                LSize = GlobalSize(hMem)
                lpMem = GlobalLock(hMem)
                ReDim abData(LSize - 1)
                CopyMemory abData(0), ByVal lpMem, LSize
                GlobalUnlock hMem
                ArrayFromPicture_O = abData
            End If

            Set IStm = Nothing
        End If

        GdipDisposeImage_O lBitmap
    End If

End Function

Public Function AttachmentToPicture_O(strTable As String, strAttachmentField As String, strImage As String) As StdPicture
    Dim strSQL As String
    Dim bin() As Byte
    Dim nOffset As Long
    Dim nSize As Long

    strSQL = "SELECT " & strTable & "." & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & _
             " WHERE " & strTable & "." & strAttachmentField & ".FileName='" & strImage & "'"
    On Error Resume Next
    bin = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenSnapshot)(0)
    If Err.Number = 0 Then
        Dim bin2() As Byte
        nOffset = bin(0)
        nSize = UBound(bin)
        ReDim bin2(nSize - nOffset)
        CopyMemory bin2(0), bin(nOffset), nSize - nOffset
        Set AttachmentToPicture_O = ArrayToPicture_O(bin2)
        Erase bin2
        Erase bin
    End If
End Function

Public Function ArrayToPicture_O(ByRef PicBin() As Byte, Optional FlattenColor As Variant) As StdPicture
    Dim IStm As IUnknown
    Dim lBitmap As Long
    Dim hBmp As Long
    Dim ret As Long

    If Not InitGDIP Then Exit Function

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)
    If ret = 0 Then
        ret = GdipLoadImageFromStream_O(IStm, lBitmap)
        If ret = 0 Then
            If Not IsMissing(FlattenColor) Then
                Dim lBitmap2 As Long
                Dim lGraph As Long
                Dim W As Long, H As Long

                ret = GdipCloneImage_O(lBitmap, lBitmap2)
                ret = GdipGetImageGraphicsContext_O(lBitmap2, lGraph)
                If ret = 0 Then
                    ret = GdipGetImageWidth_O(lBitmap, W)
                    ret = GdipGetImageHeight_O(lBitmap, H)
                    ret = GdipGraphicsClear_O(lGraph, CLng(FlattenColor))
                    ret = GdipDrawImageRectRectI_O(lGraph, lBitmap, 0, 0, W, H, 0, 0, W, H, _
                                                 UnitPixel, 0, 0)
                End If
                GdipCreateHBITMAPFromBitmap_O lBitmap2, hBmp, 0&
                GdipDeleteGraphics_O lGraph
            Else
                GdipCreateHBITMAPFromBitmap_O lBitmap, hBmp, 0&
            End If
            If hBmp <> 0 Then
                Set ArrayToPicture_O = BitmapToPicture(hBmp)
            End If
        End If
        GdipDisposeImage_O lBitmap
    End If

End Function


Function MaskFromPicture_O(ByVal Image As StdPicture, Optional TransColor As Variant) As StdPicture
    Dim lBitmap As Long
    Dim hBitmap As Long
    Dim W As Long, H As Long
    Dim bytes() As Long
    Dim BD As BitmapData
    Dim rct As RECTL
    Dim X As Long, Y As Long
    Dim AlphaColor As Long
    Dim ret As Long

    If Not InitGDIP Then Exit Function

    ret = GdipCreateBitmapFromHBITMAP_O(Image.Handle, 0, lBitmap)
    If ret = 0 Then
        ret = GdipGetImageWidth_O(lBitmap, W)
        ret = GdipGetImageHeight_O(lBitmap, H)
        With rct
            .Left = 0
            .Top = H
            .Right = W
            .Bottom = 0
        End With
        ReDim bytes(W, H)
        With BD
            .Width = W
            .Height = H
            .PixelFormat = PixelFormat32bppARGB
            .stride = 4 * CLng(.Width + 1)
            .scan0 = VarPtr(bytes(0, 0))
        End With
        ret = GdipBitmapLockBits_O(lBitmap, rct, ImageLockModeRead Or _
                                               ImageLockModeUserInputBuf Or ImageLockModewrite, PixelFormat32bppARGB, BD)
        If IsMissing(TransColor) Then
            AlphaColor = bytes(0, 0)
        Else
            AlphaColor = CLng(TransColor)
        End If
        For X = 0 To W
            For Y = 0 To H
                If bytes(X, Y) = AlphaColor Then bytes(X, Y) = &HFFFFFF Else bytes(X, Y) = &H0
            Next Y
        Next X

        ret = GdipBitmapUnlockBits_O(lBitmap, BD)
        GdipCreateHBITMAPFromBitmap_O lBitmap, hBitmap, 0&
        Set MaskFromPicture_O = BitmapToPicture(hBitmap)
        GdipDisposeImage_O lBitmap
    End If

End Function