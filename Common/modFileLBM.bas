Attribute VB_Name = "modFileLBM"
Option Explicit
Public Function IsLBM(FileName As String) As Byte

' Purpose: See if the file is a valid LBM file.

On Error GoTo FileError

Dim sHeader As String * 4

Open FileName For Binary Access Read As FILENUM_PICTURE

'Is this real a DPaint pic ?
Get FILENUM_PICTURE, , sHeader

If sHeader = "FORM" Then IsLBM = True

Close FILENUM_PICTURE

Exit Function

FileError:
ErrHandle
Resume Next

End Function
Public Sub LoadLBM(FileName As String, pImage As IMAGEFILE)

On Error GoTo FileError

Dim nIndex As Long
Dim lTemp As Long
Dim sHeader As String * 4
Dim lPictureWidth As Long
Dim lPictureHeight As Long
Dim bBitsPerPixel As Byte
Dim iNumOfColours As Integer
Dim bPacked As Byte
Dim bTempByte As Byte

Dim lFilePointer As Long
Dim bFileContainer() As Byte
Dim lOffset As Long

Dim bCounterA As Byte
Dim bCounterB As Byte
Dim lPointerX As Long

Open FileName For Binary Access Read As FILENUM_PICTURE

'Find the Information Header
For nIndex = 9 To LOF(FILENUM_PICTURE)
    Get FILENUM_PICTURE, nIndex, sHeader
    If sHeader = "BMHD" Then Seek FILENUM_PICTURE, nIndex + 8: Exit For
Next nIndex

'Get Picture Information
Get FILENUM_PICTURE, , bTempByte
lPictureWidth = bTempByte * 256
Get FILENUM_PICTURE, , bTempByte
lPictureWidth = lPictureWidth + bTempByte

Get FILENUM_PICTURE, , bTempByte
lPictureHeight = bTempByte * 256
Get FILENUM_PICTURE, , bTempByte
lPictureHeight = lPictureHeight + bTempByte

Seek FILENUM_PICTURE, Seek(FILENUM_PICTURE) + 4
Get FILENUM_PICTURE, , bBitsPerPixel
If bBitsPerPixel = 1 Then iNumOfColours = 2
If bBitsPerPixel = 4 Then iNumOfColours = 16
If bBitsPerPixel = 8 Then iNumOfColours = 256

Seek FILENUM_PICTURE, Seek(FILENUM_PICTURE) + 1
Get FILENUM_PICTURE, , bPacked

If bBitsPerPixel = 8 And bPacked = 1 Then
    With pImage
        .Width = lPictureWidth
        .Height = lPictureHeight
        .BPP = bBitsPerPixel
        ReDim .Palette(0 To 255)
    End With

    Seek FILENUM_PICTURE, 1

    'get the palette (or colour map)
    For nIndex = 1 To LOF(FILENUM_PICTURE)
        Get FILENUM_PICTURE, nIndex, sHeader
        If sHeader = "CMAP" Then Seek FILENUM_PICTURE, nIndex + 8: Exit For
    Next nIndex

    For nIndex = 0 To iNumOfColours - 1
        Get FILENUM_PICTURE, , bTempByte
        lTemp = bTempByte

        Get FILENUM_PICTURE, , bTempByte
        lTemp = lTemp + CLng(bTempByte) * 256

        Get FILENUM_PICTURE, , bTempByte
        lTemp = lTemp + CLng(bTempByte) * 65536

        With pImage
            .Palette(nIndex).rgbRed = Int(lTemp Mod 256)
            .Palette(nIndex).rgbGreen = Int(lTemp / 256) Mod 256
            .Palette(nIndex).rgbBlue = Int(lTemp / 65536)
        End With
    Next nIndex

    'Find where the picture data starts.
    For nIndex = Seek(FILENUM_PICTURE) To (LOF(FILENUM_PICTURE) - 4)
        Get FILENUM_PICTURE, nIndex, sHeader
        If sHeader = "BODY" Then Seek FILENUM_PICTURE, nIndex + 8: Exit For
    Next nIndex

    ReDim bFileContainer(1 To (LOF(FILENUM_PICTURE) - Loc(FILENUM_PICTURE)))
    lFilePointer = 1

    Get FILENUM_PICTURE, Seek(FILENUM_PICTURE), bFileContainer
    'Decompress picture data
    ReDim pImage.Data(1 To (pImage.Width * pImage.Height))
    lOffset = 1

    Do Until lFilePointer >= UBound(bFileContainer)
        bCounterA = bFileContainer(lFilePointer)

        lFilePointer = lFilePointer + 1

        If bCounterA > 128 Then
            bCounterB = bFileContainer(lFilePointer)
            lFilePointer = lFilePointer + 1
            For nIndex = lPointerX To ((lPointerX + (257 - bCounterA)) - 1)
                pImage.Data(lOffset) = bCounterB
                lOffset = lOffset + 1
            Next nIndex
            lPointerX = lPointerX + (257 - bCounterA)
            If lPointerX > lPictureWidth - 1 Then lPointerX = 0
        Else
            For nIndex = 0 To bCounterA
                bCounterB = bFileContainer(lFilePointer)
                lFilePointer = lFilePointer + 1
                pImage.Data(lOffset) = bCounterB
                lOffset = lOffset + 1
                lPointerX = 0
                If lPointerX > lPictureWidth - 1 Then lPointerX = 0
            Next nIndex
        End If
    Loop

    Erase bFileContainer
End If

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
