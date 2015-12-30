Attribute VB_Name = "modFilePCX"
Option Explicit

' Palette type.
Private Type IMyRGB
    r As Byte
    g As Byte
    b As Byte
End Type

' PCX  header
Private Type PCXHeader
    PCXFlag As Byte
    PCXVersion As Byte
    RunLengthEncode As Byte
    BitsPerPixel As Byte
    XStart As Integer
    YStart As Integer
    XEnd As Integer
    YEnd As Integer
    HorResolution As Integer
    VerResolution As Integer
    ColorMap(0 To 15) As IMyRGB
    Reserved As Byte
    NumPlanes As Byte
    BytesPerLine As Integer
    PaletteInterp As Integer
    AlsoReserved(0 To 57) As Byte
End Type
             
Private Type PaletteTable
    PaletteFlag As Byte
    Palette(0 To 255) As IMyRGB
End Type
Public Function IsPCX(FileName As String) As Byte

' Purpose: See if the file is a valid PCX file.

On Error GoTo FileError

Dim tPCXHeader As PCXHeader

Open FileName For Binary Access Read As FILENUM_PICTURE

Get FILENUM_PICTURE, , tPCXHeader.PCXFlag

Close FILENUM_PICTURE

If tPCXHeader.PCXFlag = 10 Then
    IsPCX = True
End If

Exit Function

FileError:
ErrHandle
Resume Next

End Function
Public Sub LoadPCX(FileName As String, ByRef pImage As IMAGEFILE)

On Error GoTo FileError

Dim lBytesPerLine As Long
Dim nIndex As Long
Dim nIndex2 As Long
Dim lOffset As Long
Dim lCount As Long
Dim lOffset2 As Long
Dim lRun As Long, lCounter As Long
Dim tPCXHeader As PCXHeader
Dim bPalette(0 To 768) As Byte
Dim bBuffer() As Byte
Dim bDecodePlanes() As Byte

Open FileName For Binary Access Read As FILENUM_PICTURE

Get FILENUM_PICTURE, , tPCXHeader.PCXFlag
Get FILENUM_PICTURE, , tPCXHeader.PCXVersion
Get FILENUM_PICTURE, , tPCXHeader.RunLengthEncode
Get FILENUM_PICTURE, , tPCXHeader.BitsPerPixel

With tPCXHeader
    .XStart = fGetWord(FILENUM_PICTURE, True)
    .YStart = fGetWord(FILENUM_PICTURE, True)
    .XEnd = fGetWord(FILENUM_PICTURE, True)
    .YEnd = fGetWord(FILENUM_PICTURE, True)
    .HorResolution = fGetWord(FILENUM_PICTURE, True)
    .VerResolution = fGetWord(FILENUM_PICTURE, True)

    Get FILENUM_PICTURE, , .ColorMap
    Get FILENUM_PICTURE, , .Reserved
    Get FILENUM_PICTURE, , .NumPlanes

    .BytesPerLine = fGetWord(FILENUM_PICTURE, True)
    .PaletteInterp = fGetWord(FILENUM_PICTURE, True)

    Get FILENUM_PICTURE, , .AlsoReserved
    If tPCXHeader.PCXFlag <> 10 Then
        'invalid PCX file
        Close FILENUM_PICTURE
        Exit Sub
    End If
End With

With pImage
    .Width = tPCXHeader.XEnd - tPCXHeader.XStart + 1
    .Height = tPCXHeader.YEnd - tPCXHeader.YStart + 1
    .BPP = tPCXHeader.NumPlanes * 8
    Erase .Palette
End With

'Get Data
ReDim bBuffer(LOF(FILENUM_PICTURE) - Len(tPCXHeader))
Get FILENUM_PICTURE, , bBuffer

If pImage.BPP = 8 Then
    ' Get Palette
    Seek FILENUM_PICTURE, LOF(FILENUM_PICTURE) - 768
    Get FILENUM_PICTURE, , bPalette
    If bPalette(0) <> 12 Then
        'invalid palette
        Close FILENUM_PICTURE
        Exit Sub
    End If

    'fill out the palette
    ReDim pImage.Palette(0 To 255)
    For nIndex = 0 To 255
        pImage.Palette(nIndex).rgbRed = bPalette((nIndex * 3) + 1)
        pImage.Palette(nIndex).rgbGreen = bPalette((nIndex * 3) + 2)
        pImage.Palette(nIndex).rgbBlue = bPalette((nIndex * 3) + 3)
    Next nIndex
End If

Close FILENUM_PICTURE

lBytesPerLine = tPCXHeader.NumPlanes * tPCXHeader.BytesPerLine

If pImage.BPP = 8 Then
    ReDim pImage.Data(1 To pImage.Width * pImage.Height)

    'Decompress the PCX file
    For nIndex = 0 To pImage.Height - 1
        Do Until lCount >= lBytesPerLine
            If (bBuffer(lOffset) And &HC0) = &HC0 Then
                lRun = bBuffer(lOffset) And &H3F

                If lRun + lCount > lBytesPerLine Then lRun = pImage.Width - lCount

                'Repeat
                For lCounter = 0 To lRun - 1
                    pImage.Data(((nIndex * pImage.Width) + (lCount + lCounter)) + 1) = bBuffer(lOffset + 1)
                Next lCounter

                lCount = lCount + lCounter

                'Increase our data lOffset.
                lOffset = lOffset + 2
            Else
                'If this isn't a 'lCounter' byte
                'Put the data straight into our image
                pImage.Data(((nIndex * pImage.Width) + lCount) + 1) = bBuffer(lOffset)

                lCount = lCount + 1

                'And increase the data lCount by 1
                lOffset = lOffset + 1
            End If
        Loop

        lCount = 0
    Next nIndex
Else
    nIndex2 = 0

    ReDim pImage.Data(1 To (pImage.Width * pImage.Height * 3&))
    ReDim bDecodePlanes(pImage.Width * pImage.Height * 3&)

    For nIndex = 0 To UBound(bBuffer)
        If (bBuffer(nIndex) And &HC0) = &HC0 Then
            'Encoded data
            lRun = bBuffer(nIndex) And &H3F
            
            For lCount = 0 To lRun - 1
                bDecodePlanes(nIndex2 + lCount) = bBuffer(nIndex + 1)
            Next lCount
            
            nIndex = nIndex + 1 ' Double incrememnt due to count byte
            nIndex2 = nIndex2 + lRun 'Indicate how many pixels added
        Else
            'Not Encoded
            bDecodePlanes(nIndex2) = bBuffer(nIndex)

            nIndex2 = nIndex2 + 1
        End If
    Next nIndex

    lOffset2 = 1

    For nIndex = 0 To pImage.Height - 1
        For nIndex2 = 0 To pImage.Width - 1
            lOffset = nIndex * tPCXHeader.BytesPerLine * 3 + nIndex2
            pImage.Data(lOffset2) = bDecodePlanes(lOffset + tPCXHeader.BytesPerLine * 2)
            pImage.Data(lOffset2 + 1) = bDecodePlanes(lOffset + tPCXHeader.BytesPerLine)
            pImage.Data(lOffset2 + 2) = bDecodePlanes(lOffset)
            lOffset2 = lOffset2 + 3
        Next nIndex2
    Next nIndex

    Erase bDecodePlanes
End If

Erase bBuffer

Exit Sub

FileError:
ErrHandle
On Error GoTo 0
Close FILENUM_PICTURE

' Clear any information already in pImage
Dim pBlank As IMAGEFILE
pImage = pBlank

End Sub
