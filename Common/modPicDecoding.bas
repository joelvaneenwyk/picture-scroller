Attribute VB_Name = "modPicDecoding"
Option Explicit

Public Const PIC_PCX = 7
Public Const PIC_PSD = 8
Public Const PIC_TGA = 9
Public Const PIC_LBM = 10

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type IMAGEFILE
    Width As Long
    Height As Long
    BPP As Byte
    Palette() As RGBQUAD
    Data() As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal DX As Long, ByVal DY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long

On Error GoTo ErrorHandler

LShift = lValue * (2 ^ lNumberOfBitsToShift)

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function fGetWord(iFileNumber As Integer, bIntel As Boolean) As Long

On Error GoTo ErrorHandler

Dim lFigure1 As Long, lFigure2 As Long, bByte1 As Byte, bByte2 As Byte

Get #iFileNumber, , bByte1
Get #iFileNumber, , bByte2

lFigure1 = bByte1 And 255
lFigure2 = bByte2 And 255

If bIntel = True Then
    fGetWord = (lFigure1 + (LShift(lFigure2, 8)))
Else
    fGetWord = (LShift(lFigure1, 8) + lFigure2)
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub DrawImage(hdc As Long, ByRef pImage As IMAGEFILE, ByVal iWidth As Integer, ByVal iHeight As Integer)

On Error GoTo ErrorHandler

Dim tBitmap As BITMAPINFO

With tBitmap.bmiHeader
    'Simply the size of the header type
    .biSize = 40

    'Width of Image
    .biWidth = pImage.Width

    'Height of image (reversed)
    .biHeight = -pImage.Height

    'The image has 1 plane
    .biPlanes = 1

    'How many Bits Per Pixel.
    .biBitCount = pImage.BPP

    'No Compression used.
    .biCompression = 0
End With

If pImage.BPP = 8 Then CopyMemory tBitmap.bmiColors(0), pImage.Palette(0), 256 * 4

'Now take this information and copy the data onto the main pictures HDC.
'Remember that StretchDIBits also draws from the bottom left corner upwards, so if our
'data is stored from top to bottom, then we must negate the height value which will
'flip the image.
StretchDIBits hdc, 0, 0, iWidth, iHeight, 0, 0, pImage.Width, pImage.Height, pImage.Data(1), tBitmap, 0, vbSrcCopy

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
