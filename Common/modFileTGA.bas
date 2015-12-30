Attribute VB_Name = "modFileTGA"
Option Explicit

' Tagar file header
Private Type TRGHDR
    'Characters in ID field
    bIDFieldSize As Byte
    'Color map type
    bClrMapType As Byte
    'Image type
    bImageType As Byte
    'Color map specification
    lClrMapSpec(0 To 4) As Byte
    'X origin
    wXOrigin As Integer
    'Y origin
    wYOrigin As Integer
    'Bitmap width
    wWidth As Long
    'Bitmap height
    wHeight As Long
    'Bits per pixel
    bBitsPixel As Byte
    'Image descriptor
    bImageDescriptor As Byte
End Type
Public Function IsTGA(FileName As String) As Byte

' Purpose: See if the file is a valid TGA file.

On Error GoTo FileError

Dim iTemp As Integer
Dim tHeader As TRGHDR

Open FileName For Binary Access Read As FILENUM_PICTURE

Get FILENUM_PICTURE, , tHeader.bIDFieldSize
Get FILENUM_PICTURE, , tHeader.bClrMapType
Get FILENUM_PICTURE, , tHeader.bImageType
Get FILENUM_PICTURE, , tHeader.lClrMapSpec
Get FILENUM_PICTURE, , tHeader.wXOrigin
Get FILENUM_PICTURE, , tHeader.wYOrigin
Get FILENUM_PICTURE, , iTemp
Get FILENUM_PICTURE, , iTemp
Get FILENUM_PICTURE, , tHeader.bBitsPixel
Get FILENUM_PICTURE, , tHeader.bImageDescriptor

If tHeader.bImageType = 2 And tHeader.bBitsPixel = 24 Then IsTGA = True

Close FILENUM_PICTURE

Exit Function

FileError:
ErrHandle
Resume Next

End Function
Public Sub LoadTGA(FileName As String, ByRef pImage As IMAGEFILE)

On Error GoTo FileError

Dim tHeader As TRGHDR
Dim iTemp As Integer
Dim bBuffer() As Byte
Dim nIndex As Long
Dim lOffset1 As Long

Open FileName For Binary Access Read As FILENUM_PICTURE

Get FILENUM_PICTURE, , tHeader.bIDFieldSize
Get FILENUM_PICTURE, , tHeader.bClrMapType
Get FILENUM_PICTURE, , tHeader.bImageType
Get FILENUM_PICTURE, , tHeader.lClrMapSpec
Get FILENUM_PICTURE, , tHeader.wXOrigin
Get FILENUM_PICTURE, , tHeader.wYOrigin
Get FILENUM_PICTURE, , iTemp
tHeader.wWidth = iTemp

Get FILENUM_PICTURE, , iTemp
tHeader.wHeight = iTemp

Get FILENUM_PICTURE, , tHeader.bBitsPixel
Get FILENUM_PICTURE, , tHeader.bImageDescriptor

If (tHeader.bImageType <> 2) Or (tHeader.bBitsPixel <> 24) Then Close #1: Exit Sub   'not an TGA 24 bits file

With pImage
    .BPP = 24
    .Width = tHeader.wWidth
    .Height = tHeader.wHeight
    Erase .Palette
    ReDim .Data(1 To (CLng(tHeader.wWidth) * 3) * tHeader.wHeight) As Byte
    ReDim bBuffer(1 To tHeader.wWidth * 3)

    lOffset1 = UBound(.Data)
    lOffset1 = lOffset1 - ((.Width * 3) - 1)

    For nIndex = tHeader.wHeight - 1 To 0 Step -1
        Get FILENUM_PICTURE, , bBuffer
        CopyMemory .Data(lOffset1), bBuffer(1), (.Width * 3)
        lOffset1 = lOffset1 - (.Width * 3)
    Next nIndex
End With

Close FILENUM_PICTURE

Exit Sub

FileError:
ErrHandle
On Error GoTo 0
Close FILENUM_PICTURE

' Clear any information already in pImage
Dim pBlank As IMAGEFILE
pImage = pBlank

End Sub
