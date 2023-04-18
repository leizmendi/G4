Attribute VB_Name = "basQRcode"
' $Id: basQRcode.bas
'/**
' The VBA/VB6 interface to diQRcode.dll.
'
' @author dai
' @version 4.0.0
'**/
' $Date: 2021-06-16 10:07:00 $

'*********************** COPYRIGHT NOTICE***********************
' Copyright (c) 2021 DI Management Services Pty Limited.
' <www.di-mgt.com.au> <www.cryptosys.net>. All rights reserved.
' This code may only be used by licensed users and in
' accordance with the licence conditions.
' The latest version of diQRcode and a licence
' may be obtained from <https://www.cryptosys.net/qrcode>.
' This copyright notice must always be left intact.
'******************** END OF COPYRIGHT NOTICE*******************

Option Explicit
Option Base 0

'//////////////////////////////////////////////////
' ECC LEVELS
Public Const QRCODE_ECC_M As Long = 0  ' Default
Public Const QRCODE_ECC_L As Long = 1
Public Const QRCODE_ECC_Q As Long = 2
Public Const QRCODE_ECC_H As Long = 3
' OPTION FLAGS
Public Const QRCODE_ESCAPED As Long = &H1000&
Public Const QRCODE_BASE64  As Long = &H2000&
Public Const QRCODE_NO_NAMECHECK As Long = &H10000

'/**
' Image type
'**/
Public Enum ImageType
    '/**
    ' Output image as a GIF file (default)
    '**/
    GIF = 0
    '/**
    ' Output image as an SVG file
    '**/
    SVG = &H10
End Enum

'/**
' Error-correction code
'**/
Public Enum Ecc
    '/**
    ' EC level M (default)
    '**/
    M = 0
    '/**
    ' EC level L
    '**/
    L = 1
    '/**
    ' EC level Q
    '**/
    Q = 2
    '/**
    ' EC level H
    '**/
    H = 3
End Enum

' Local constants
Private Const OUT_OF_RANGE_ERROR As Long = 11

#If VBA7 Then
' Declarations for 64-bit Office
' (In VB6 these will appear red. Turn off "Auto Syntax Check" in Tools > Options to avoid annoying warnings)
Private Declare PtrSafe Function QRCODE_CreateImage Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
Private Declare PtrSafe Function QRCODE_CreateImageFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateImage" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
Private Declare PtrSafe Function QRCODE_CreateGif Lib "diQRcode.dll" (ByVal strOutputFile As String, ByVal strInput As String, ByVal nPixelsPerModule As Long, ByVal strParams As String, ByVal nOptions As Long) As Long
Private Declare PtrSafe Function QRCODE_CreateGifFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateGif" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
Private Declare PtrSafe Function QRCODE_CreatePdf Lib "diQRcode.dll" (ByVal strOutputFile As String, ByVal strInput As String, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
Private Declare PtrSafe Function QRCODE_CreatePdfFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreatePdf" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
Public Declare PtrSafe Function QRCODE_Version Lib "diQRcode.dll" () As Long
Private Declare PtrSafe Function QRCODE_DllInfo Lib "diQRcode.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal nOptions As Long) As Long
Private Declare PtrSafe Function QRCODE_ErrorLookup Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nErrCode As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#Else
' Declarations for VB6 and 32-bit Office
Private Declare Function QRCODE_CreateImage Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreateImageFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateImage" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nMargin As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreateGif Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreateGifFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreateGif" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal szParams As String, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreatePdf Lib "diQRcode.dll" (ByVal szOutputFile As String, ByVal szInput As String, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_CreatePdfFromBytes Lib "diQRcode.dll" Alias "QRCODE_CreatePdf" (ByVal szOutputFile As String, ByRef lpInput As Byte, ByVal nPixelsPerModule As Long, ByVal nPageWidth As Long, ByVal nPageHeight As Long, ByVal nX As Long, ByVal nY As Long, ByVal nOptions As Long) As Long
Private Declare Function QRCODE_ErrorLookup Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nErrCode As Long) As Long

'/**
' Get version number of core DLL.
' @return Version number as an integer in form `Major*10000 + Minor*100 + Release`. For example, version 2.10.3 would return 21003.
'**/
Public Declare Function QRCODE_Version Lib "diQRcode.dll" () As Long
Private Declare Function QRCODE_DllInfo Lib "diQRcode.dll" (ByVal szOutput As String, ByVal nOutChars As Long, ByVal nOptions As Long) As Long

'///////////////////////////////////////////////////////////////
' INTERNAL UTF-8 MAPPING
''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#End If
    
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
' Ref: [How to convert VBA/VB6 Unicode strings to UTF-8](https://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html)
Private Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

' ------------------
' EXPORTED FUNCTIONS
' ------------------

'/**
' Create an image file of a QR code.
' @param szOutputFile Name of output image file to be created.
' @param szInput Text input to be encoded (ANSI characters only).
' @param imgType Image type (default = GIF)
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param nMargin Size of margin in modules (default = 4 modules)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nOptions Option flags. Use the `Or` operator to combine: <br>
' `QRCODE_ESCAPED` to indicate #-escaped sequences in string <br>
' `QRCODE_BASE64` to encode output as base64 text <br>
' `QRCODE_NO_NAMECHECK` do not check filename extension against file type <br>
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'**/
Public Function qrcodeCreateImage(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.M, Optional nOptions As Long = 0) As Long
    Dim i As Long
    Dim opts As Long
    opts = nOptions Or EccLevel Or imgType
    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
    For i = 1 To Len(szInput)
        If AscW(Mid(szInput, i, 1)) > &HFF Then
            qrcodeCreateImage = -OUT_OF_RANGE_ERROR
            Exit Function
        End If
    Next
    ' Do the business
    qrcodeCreateImage = QRCODE_CreateImage(szOutputFile, szInput, nPixelsPerModule, nMargin, opts)
End Function

'/**
' Create a GIF file of a QR code (_deprecated_).
' @param szOutputFile Name of output GIF file to be created.
' @param szInput Text input to be encoded (ANSI characters only).
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param szParams Optional parameters.  Set as `"margin=N"` to change the margin to `N` modules (default=4).
' @param nOptions Option flags. Set as `QRCODE_ESCAPED` to indicate #-escaped sequences in string.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
' @deprecated Use {@link qrcodeCreateImage}.
'**/
Public Function qrcodeCreateGif(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.M, Optional szParams As String = "", Optional nOptions As Long = 0) As Long
    Dim i As Long
    Dim opts As Long
    Const OUT_OF_RANGE_ERROR As Long = 11
    opts = nOptions Or EccLevel
    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
    For i = 1 To Len(szInput)
        If AscW(Mid(szInput, i, 1)) > &HFF Then
            qrcodeCreateGif = -OUT_OF_RANGE_ERROR
            Exit Function
        End If
    Next
    ' Do the business
    qrcodeCreateGif = QRCODE_CreateGif(szOutputFile, szInput, nPixelsPerModule, szParams, opts)
End Function

'/**
' Create a PDF file of a QR code.
' @param szOutputFile Name of output PDF file to be created.
' @param szInput Text input to be encoded (ANSI characters only).
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nPageWidth  Width of PDF page in pixels (default = 0 : set width to fit QRcode image).
' @param nPageHeight Height of PDF page in pixels (default = 0 : set height to fit QRcode image).
' @param nX  X-coordinate in pixels of bottom-left of QRcode image (default = 0 : at left side).
' @param nY  Y-coordinate in pixels of bottom-left of QRcode image (default = 0 : at bottom).
' @param nOptions Option flags. Set as `QRCODE_ESCAPED` to indicate #-escaped sequences in string.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'**/
Public Function qrcodeCreatePdf(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.M, Optional nPageWidth As Long = 0, Optional nPageHeight As Long = 0, Optional nX As Long = 0, Optional nY As Long = 0, Optional nOptions As Long = 0) As Long
    Dim i As Long
    Dim opts As Long
    Const OUT_OF_RANGE_ERROR As Long = 11
    opts = nOptions Or EccLevel
    ' Input text must consist of ANSI chars only, else OUT_OF_RANGE_ERROR.
    For i = 1 To Len(szInput)
        If AscW(Mid(szInput, i, 1)) > &HFF Then
            qrcodeCreatePdf = -OUT_OF_RANGE_ERROR
            Exit Function
        End If
    Next
    ' Do the business
    qrcodeCreatePdf = QRCODE_CreatePdf(szOutputFile, szInput, nPixelsPerModule, nPageWidth, nPageHeight, nX, nY, opts)
End Function

'/**
' Create a image file of a QR code, encoding input in UTF-8.
' @param szOutputFile Name of output image file to be created.
' @param szInput Text input to be encoded.
' @param imgType Image type (default = GIF)
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param nMargin Size of margin in modules (default = 4 modules)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nOptions Option flags. Use the `Or` operator to combine: <br>
' `QRCODE_ESCAPED` to indicate #-escaped sequences in string <br>
' `QRCODE_BASE64` to encode output as base64 text <br>
' `QRCODE_NO_NAMECHECK` do not check filename extension against file type <br>
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check)
'**/
Public Function qrcodeCreateImageInUtf8(szOutputFile As String, szInput As String, Optional imgType As ImageType = ImageType.GIF, _
    Optional nPixelsPerModule As Long = 0, Optional nMargin As Long = 0, Optional EccLevel As Ecc = Ecc.M, Optional nOptions As Long = 0) As Long
    Dim b() As Byte
    Dim nLen As Long
    Dim opts As Long
    opts = nOptions Or EccLevel Or imgType
    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
    b = Utf8BytesFromString(szInput)
    ' Add an extra NUL byte to the array
    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
    ReDim Preserve b(nLen)
    b(nLen - 1) = 0
    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
    qrcodeCreateImageInUtf8 = QRCODE_CreateImageFromBytes(szOutputFile, b(0), nPixelsPerModule, nMargin, opts)
End Function

'/**
' Create a GIF file of a QR code, encoding input in UTF-8 (_deprecated_).
' @param szOutputFile Name of output GIF file to be created.
' @param szInput Text input to be encoded.
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param szParams Optional parameters. Set as `"margin=N"` to change the margin to `N` modules (default=4).
' @param nOptions Option flags. Set as 0 for default.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check).
' @remark Any non-ASCII characters in `szInput` will be encoded in UTF-8 before processing.
' @deprecated Use {@link qrcodeCreateImageInUtf8}.
'**/
Public Function qrcodeCreateGifInUtf8(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.M, Optional szParams As String = "", Optional nOptions As Long = 0) As Long
    Dim b() As Byte
    Dim nLen As Long
    Dim opts As Long
    opts = nOptions Or EccLevel
    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
    b = Utf8BytesFromString(szInput)
    ' Add an extra NUL byte to the array
    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
    ReDim Preserve b(nLen)
    b(nLen - 1) = 0
    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
    qrcodeCreateGifInUtf8 = QRCODE_CreateGifFromBytes(szOutputFile, b(0), nPixelsPerModule, szParams, opts)
End Function


'/**
' Create a PDF file of a QR code, encoding input in UTF-8.
' @param szOutputFile Name of output GIF file to be created.
' @param szInput Text input to be encoded.
' @param nPixelsPerModule Number of pixels per module (default = 2 ppm)
' @param EccLevel Error correction level (default = Ecc.M)
' @param nPageWidth  Width of PDF page in pixels (default = 0 : set width to fit QRcode image).
' @param nPageHeight Height of PDF page in pixels (default = 0 : set height to fit QRcode image).
' @param nX  X-coordinate in pixels of bottom-left of QRcode image (default = 0 : at left side).
' @param nY  Y-coordinate in pixels of bottom-left of QRcode image (default = 0 : at bottom).
' @param nOptions Option flags. Set as 0 for default.
' @return Zero on success, or a nonzero error code (use {@link qrcodeErrorLookup} to check).
' @remark Any non-ASCII characters in `szInput` will be encoded in UTF-8 before processing.
'**/
Public Function qrcodeCreatePdfInUtf8(szOutputFile As String, szInput As String, Optional nPixelsPerModule As Long = 0, Optional EccLevel As Ecc = Ecc.M, Optional nPageWidth As Long = 0, Optional nPageHeight As Long = 0, Optional nX As Long = 0, Optional nY As Long = 0, Optional nOptions As Long = 0) As Long
    Dim b() As Byte
    Dim nLen As Long
    Dim opts As Long
    opts = nOptions Or EccLevel
    ' Convert input Unicode (UTF-16) string to UTF-8-encoded sequence of bytes
    b = Utf8BytesFromString(szInput)
    ' Add an extra NUL byte to the array
    nLen = UBound(b) + 2    ' b is currently Ubound(b) - 1 bytes long
    ReDim Preserve b(nLen)
    b(nLen - 1) = 0
    ' Pass nul-terminated UTF-8 bytes to aliased form of core function
    qrcodeCreatePdfInUtf8 = QRCODE_CreatePdfFromBytes(szOutputFile, b(0), nPixelsPerModule, nPageWidth, nPageHeight, nX, nY, opts)
End Function

'/**
' Get information about the core native DLL.
' @param  nOptions For future use.
' @return Information about the core DLL module.
' @remark Example result.
' `Platform` is the platform the core DLL was compiled for: `Win32` or `X64`.
' {@code
' "Platform=Win32; Compiled=Mar 27 2021 03:50:03; Licence=T"
' }
'**/
Public Function qrcodeDllInfo(Optional nOptions As Long = 0) As String
    Dim nc As Long
    nc = QRCODE_DllInfo("", 0, nOptions)
    If nc <= 0 Then Exit Function
    qrcodeDllInfo = String(nc, " ")
    nc = QRCODE_DllInfo(qrcodeDllInfo, Len(qrcodeDllInfo), nOptions)
    qrcodeDllInfo = Left$(qrcodeDllInfo, nc)
End Function

'/**
' Look up description for error code.
' @param  nErrCode Value of error code to lookup (may be positive or negative).
' @return Error message, or empty string if no corresponding error code.
'**/
Public Function qrcodeErrorLookup(nErrCode As Long) As String
    Dim nc As Long
    nc = QRCODE_ErrorLookup("", 0, nErrCode)
    If nc <= 0 Then Exit Function
    qrcodeErrorLookup = String(nc, " ")
    nc = QRCODE_ErrorLookup(qrcodeErrorLookup, nc, nErrCode)
    qrcodeErrorLookup = Left$(qrcodeErrorLookup, nc)
End Function
