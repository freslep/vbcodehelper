Attribute VB_Name = "MButtonManipulation"
''*******************************************************************************
'' MODULE:       MButtonManipulation
'' FILENAME:     C:\My Code\vb\vbch\Source\MButtonManipulation.bas
'' AUTHOR:       Phil Fresle
'' CREATED:      11-Nov-2001
'' COPYRIGHT:    Copyright 2002-2019 Frez Systems Limited.
''
'' DESCRIPTION:
'' ***Description goes here***
''
'' MODIFICATION HISTORY:
'' 1.0       06-Mar-2002
''           Phil Fresle
''           Initial Version
''*******************************************************************************
'Option Explicit
'
'Public Type BITMAPINFOHEADER '40 bytes
'   biSize As Long
'   biWidth As Long
'   biHeight As Long
'   biPlanes As Integer
'   biBitCount As Integer
'   biCompression As Long
'   biSizeImage As Long
'   biXPelsPerMeter As Long
'   biYPelsPerMeter As Long
'   biClrUsed As Long
'   biClrImportant As Long
'End Type
'
'Public Type BITMAP
'   bmType As Long
'   bmWidth As Long
'   bmHeight As Long
'   bmWidthBytes As Long
'   bmPlanes As Integer
'   bmBitsPixel As Integer
'   bmBits As Long
'End Type
'
'' ===================================================================
''   GDI/Drawing Functions (to build the mask)
'' ===================================================================
'Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" _
'  (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" _
'  (ByVal hdc As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
'  (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function CreateBitmap Lib "gdi32" _
'  (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, _
'   ByVal nBitCount As Long, lpBits As Any) As Long
'Private Declare Function SelectObject Lib "gdi32" _
'  (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" _
'  (ByVal hObject As Long) As Long
'Private Declare Function GetBkColor Lib "gdi32" _
'  (ByVal hdc As Long) As Long
'Private Declare Function SetBkColor Lib "gdi32" _
'  (ByVal hdc As Long, ByVal crColor As Long) As Long
'Private Declare Function GetTextColor Lib "gdi32" _
'  (ByVal hdc As Long) As Long
'Private Declare Function SetTextColor Lib "gdi32" _
'  (ByVal hdc As Long, ByVal crColor As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" _
'  (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
'   ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
'   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function CreateHalftonePalette Lib "gdi32" _
'  (ByVal hdc As Long) As Long
'Private Declare Function SelectPalette Lib "gdi32" _
'  (ByVal hdc As Long, ByVal hPalette As Long, _
'   ByVal bForceBackground As Long) As Long
'Private Declare Function RealizePalette Lib "gdi32" _
'  (ByVal hdc As Long) As Long
'Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
'  (ByVal lOleColor As Long, ByVal lHPalette As Long, _
'   lColorRef As Long) As Long
'Private Declare Function GetDIBits Lib "gdi32" _
'  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
'   ByVal nNumScans As Long, lpBits As Any, lpBI As Any, _
'   ByVal wUsage As Long) As Long
'Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
'  (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'
'' ===================================================================
''   Clipboard APIs
'' ===================================================================
'Private Declare Function OpenClipboard Lib "user32" _
'  (ByVal hwnd As Long) As Long
'Private Declare Function CloseClipboard Lib "user32" () As Long
'Private Declare Function RegisterClipboardFormat Lib "user32" _
'  Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'Private Declare Function GetClipboardData Lib "user32" _
'  (ByVal wFormat As Long) As Long
'Private Declare Function SetClipboardData Lib "user32" _
'  (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Private Declare Function EmptyClipboard Lib "user32" () As Long
'Private Const CF_DIB = 8
'
'' ===================================================================
''   Memory APIs (for clipboard transfers)
'' ===================================================================
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'  (pDest As Any, pSource As Any, ByVal cbLength As Long)
'Private Declare Function GlobalAlloc Lib "kernel32" _
'  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" _
'  (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" _
'  (ByVal hMem As Long) As Long
'Private Declare Function GlobalSize Lib "kernel32" _
'  (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" _
'  (ByVal hMem As Long) As Long
'Private Const GMEM_DDESHARE = &H2000
'Private Const GMEM_MOVEABLE = &H2
'
'' ===================================================================
''  CopyBitmapAsButtonFace
''
''  This is the public function to call to create a mask based on the
''  bitmap provided and copy both to the clipboard. The first parameter
''  is a standard VB Picture object. The second should be the color in
''  the image you want to be made transparent.
''
''  Note: This code sample does limited error handling and is designed
''  for VB only (not VBA). You will need to make changes as appropriate
''  to modify the code to suit your needs.
''
'' ===================================================================
'Public Sub CopyBitmapAsButtonFace(ByVal picSource As StdPicture, _
'                                  ByVal clrMaskColor As OLE_COLOR)
'    Dim hPal As Long
'    Dim hdcScreen As Long
''VBCH    Dim hbmButtonFace As Long
'    Dim hbmButtonMask As Long
'    Dim bDeletePal As Boolean
'    Dim lMaskClr As Long
'
'    ' Check to make sure we have a valid picture.
'    If picSource Is Nothing Then GoTo err_invalidarg
'    If picSource.Type <> vbPicTypeBitmap Then GoTo err_invalidarg
'    If picSource.Handle = 0 Then GoTo err_invalidarg
'
'    ' Get the DC for the display device we are on.
'    hdcScreen = GetDC(0)
'    hPal = picSource.hPal
'    If hPal = 0 Then
'        hPal = CreateHalftonePalette(hdcScreen)
'        bDeletePal = True
'    End If
'
'    ' Translate the OLE_COLOR value to a GDI COLORREF value based on the palette.
'    OleTranslateColor clrMaskColor, hPal, lMaskClr
'
'    ' Create a mask based on the image handed in (hbmButtonMask is the result).
'    CreateButtonMask picSource.Handle, lMaskClr, hdcScreen, _
'        hPal, hbmButtonMask
'
'    ' Let VB copy the bitmap to the clipboard (for the CF_DIB).
'    Clipboard.SetData picSource, vbCFDIB
'
'    ' Now copy the Button Mask.
'    CopyButtonMaskToClipboard hbmButtonMask, hdcScreen
'
'    ' Delete the mask and clean up (a copy is on the clipboard).
'    DeleteObject hbmButtonMask
'    If bDeletePal Then DeleteObject hPal
'    ReleaseDC 0, hdcScreen
'
'    Exit Sub
'err_invalidarg:
'    Err.Raise 481 'VB Invalid Picture Error
'End Sub
'
'' ===================================================================
''  CreateButtonMask -- Internal helper function
'' ===================================================================
'Private Sub CreateButtonMask(ByVal hbmSource As Long, _
'                             ByVal nMaskColor As Long, _
'                             ByVal hdcTarget As Long, _
'                             ByVal hPal As Long, _
'                             ByRef hbmMask As Long)
'
'    Dim hdcSource As Long
'    Dim hdcMask As Long
'    Dim hbmSourceOld As Long
'    Dim hbmMaskOld As Long
'    Dim hpalSourceOld As Long
'    Dim uBM As BITMAP
'
'    ' Get some information about the bitmap handed to us.
'    GetObjectAPI hbmSource, 24, uBM
'
'    ' Check the size of the bitmap given.
'    If uBM.bmWidth < 1 Or uBM.bmWidth > 30000 Then Exit Sub
'    If uBM.bmHeight < 1 Or uBM.bmHeight > 30000 Then Exit Sub
'
'    ' Create a compatible DC, load the palette and the bitmap.
'    hdcSource = CreateCompatibleDC(hdcTarget)
'    hpalSourceOld = SelectPalette(hdcSource, hPal, True)
'    RealizePalette hdcSource
'    hbmSourceOld = SelectObject(hdcSource, hbmSource)
'
'    ' Create a black and white mask the same size as the image.
'    hbmMask = CreateBitmap(uBM.bmWidth, uBM.bmHeight, 1, 1, ByVal 0)
'
'    ' Create a compatble DC for it and load it.
'    hdcMask = CreateCompatibleDC(hdcTarget)
'    hbmMaskOld = SelectObject(hdcMask, hbmMask)
'
'    ' All you need to do is set the mask color as the background color
'    ' on the source picture, and set the forground color to white, and
'    ' then a simple BitBlt will make the mask for you.
'    SetBkColor hdcSource, nMaskColor
'    SetTextColor hdcSource, vbWhite
'    BitBlt hdcMask, 0, 0, uBM.bmWidth, uBM.bmHeight, hdcSource, _
'        0, 0, vbSrcCopy
'
'    ' Clean up the memory DCs.
'    SelectObject hdcMask, hbmMaskOld
'    DeleteDC hdcMask
'
'    SelectObject hdcSource, hbmSourceOld
'    SelectObject hdcSource, hpalSourceOld
'    DeleteDC hdcSource
'End Sub
'
'' ===================================================================
''  CopyButtonMaskToClipboard -- Internal helper function
'' ===================================================================
'Private Sub CopyButtonMaskToClipboard(ByVal hbmMask As Long, _
'                                      ByVal hdcTarget As Long)
'    Dim cfBtnFace As Long
'    Dim cfBtnMask As Long
'    Dim hGMemFace As Long
'    Dim hGMemMask As Long
'    Dim lpData As Long
'    Dim lpData2 As Long
'    Dim hMemTmp As Long
'    Dim cbSize As Long
'    Dim arrBIHBuffer(50) As Byte
'    Dim arrBMDataBuffer() As Byte
'    Dim uBIH As BITMAPINFOHEADER
'
'    uBIH.biSize = 40
'
'    ' Get the BITMAPHEADERINFO for the mask.
'    GetDIBits hdcTarget, hbmMask, 0, 0, ByVal 0&, uBIH, 0
'    CopyMemory arrBIHBuffer(0), uBIH, 40
'
'    ' Make sure it is a mask image.
'    If uBIH.biBitCount <> 1 Then Exit Sub
'    If uBIH.biSizeImage < 1 Then Exit Sub
'
'    ' Create a temp buffer to hold the bitmap bits.
'    ReDim Preserve arrBMDataBuffer(uBIH.biSizeImage + 4) As Byte
'
'    ' Open the clipboard.
'    If Not CBool(OpenClipboard(0)) Then Exit Sub
'
'    ' Get the cf for button face and mask.
'    cfBtnFace = RegisterClipboardFormat("Toolbar Button Face")
'    cfBtnMask = RegisterClipboardFormat("Toolbar Button Mask")
'
'    ' Open DIB on the clipboard and make a copy of it for the button face.
'    hMemTmp = GetClipboardData(CF_DIB)
'    If hMemTmp <> 0 Then
'        cbSize = GlobalSize(hMemTmp)
'        hGMemFace = GlobalAlloc(&H2002, cbSize)
'        If hGMemFace <> 0 Then
'            lpData = GlobalLock(hMemTmp)
'            lpData2 = GlobalLock(hGMemFace)
'            CopyMemory ByVal lpData2, ByVal lpData, cbSize
'            GlobalUnlock hGMemFace
'            GlobalUnlock hMemTmp
'
'            If SetClipboardData(cfBtnFace, hGMemFace) = 0 Then
'                GlobalFree hGMemFace
'            End If
'
'        End If
'    End If
'
'    ' Now get the mask bits and the rest of the header.
'    GetDIBits hdcTarget, hbmMask, 0, uBIH.biSizeImage, _
'        arrBMDataBuffer(0), arrBIHBuffer(0), 0
'
'    ' Copy them to global memory and set it on the clipboard.
'    hGMemMask = GlobalAlloc(&H2002, uBIH.biSizeImage + 50)
'    If hGMemMask <> 0 Then
'        lpData = GlobalLock(hGMemMask)
'        CopyMemory ByVal lpData, arrBIHBuffer(0), 48
'        CopyMemory ByVal (lpData + 48), _
'            arrBMDataBuffer(0), uBIH.biSizeImage
'        GlobalUnlock hGMemMask
'
'        If SetClipboardData(cfBtnMask, hGMemMask) = 0 Then
'            GlobalFree hGMemMask
'        End If
'
'    End If
'
'    ' We're done.
'    CloseClipboard
'End Sub
