Attribute VB_Name = "modFilePSD"
Option Explicit

Private Type PSDInfo
    Pixels() As Byte
    BitsPerChannel As Integer

    'For Indexed or DuoTone only
    ColorData(0 To 767) As Byte
    Mode As Integer
    Width As Long
    Height As Long
    ChannelCount As Integer
    Compression As Integer
End Type

Dim lFilePointer As Long
Dim ImageInfo As PSDInfo
Private Function Read32(ByVal lFileNumber As Integer) As Long

On Error GoTo ErrorHandler

Dim bByte1 As Byte
Dim bByte2 As Byte
Dim bByte3 As Byte
Dim bByte4 As Byte

Get #lFileNumber, lFilePointer, bByte4
lFilePointer = lFilePointer + 1

Get #lFileNumber, lFilePointer, bByte3
lFilePointer = lFilePointer + 1

Get #lFileNumber, lFilePointer, bByte2
lFilePointer = lFilePointer + 1

Get #lFileNumber, lFilePointer, bByte1
lFilePointer = lFilePointer + 1

Read32 = LShift(bByte4, 24) + LShift(bByte3, 16) + LShift(bByte2, 8) + bByte1

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Function Read16(ByVal lFileNumber As Integer) As Long

On Error GoTo ErrorHandler

Dim bByte1 As Byte
Dim bByte2 As Byte

Get #lFileNumber, lFilePointer, bByte2
lFilePointer = lFilePointer + 1

Get #lFileNumber, lFilePointer, bByte1
lFilePointer = lFilePointer + 1

Read16 = bByte1 + LShift(bByte2, 8)

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function IsPSD(FileName As String) As Byte

' Purpose: See if the file is a valid PSD file.

On Error GoTo FileError

Dim lType As Long

Open FileName For Binary Access Read As FILENUM_PICTURE

lFilePointer = 1
lType = Read32(FILENUM_PICTURE)

If lType = 943870035 Then IsPSD = True

Close FILENUM_PICTURE

Exit Function

FileError:
ErrHandle
Resume Next

End Function
Public Sub LoadPSD(FileName As String, pImage As IMAGEFILE)

On Error GoTo FileError

Dim nIndex As Long
Dim nIndex2 As Long
Dim lType As Long
Dim lModeDataCount As Long
Dim lResourceDataCount As Long
Dim lReservedDataCount As Long
Dim iPSDVersion As Integer
Dim lOffset1 As Long
Dim lOffset2 As Long

' First open the file and get the important entries from the header...
Open FileName For Binary Access Read As FILENUM_PICTURE

lFilePointer = 5
iPSDVersion = Read16(FILENUM_PICTURE)

'Incorrect PSD Version, MUST be 1.
If iPSDVersion <> 1 Then Close FILENUM_PICTURE: Exit Sub

' Skip 6 Bytes, irrelevant info. Must be 0
lFilePointer = lFilePointer + 6

With ImageInfo
    .ChannelCount = Read16(FILENUM_PICTURE)

    'Incorrect Channel Count
    If .ChannelCount < 0 Or .ChannelCount > 16 Then Close FILENUM_PICTURE: Exit Sub

    .Height = Read32(FILENUM_PICTURE)
    .Width = Read32(FILENUM_PICTURE)

    'Supported values are 1,8 or 16
    .BitsPerChannel = Read16(FILENUM_PICTURE)

    'NO RGB COLOURS
    If .BitsPerChannel <> 8 Then Close FILENUM_PICTURE: Exit Sub

    ' Make sure the color mode is RGB.
    ' Supported Modes are Bitmap=0, Grayscale=1, Indexed=2,RGB=3,CMYK=4,MultiChannel=7
    ' Duotone=8,Lab=9
    .Mode = Read16(FILENUM_PICTURE)

    'ColorMode is Not RGB
    If ImageInfo.Mode <> 3 Then Close FILENUM_PICTURE: Exit Sub

    ' Skip the Mode Data. (It's the palette for indexed color; other info for other modes.)
    lModeDataCount = Read32(FILENUM_PICTURE)
    If lModeDataCount <> 0 Then lFilePointer = lFilePointer + lModeDataCount

    ' Skip the image resources. (resolution, pen tool paths, etc)
    lResourceDataCount = Read32(FILENUM_PICTURE)
    If lResourceDataCount <> 0 Then lFilePointer = lFilePointer + lResourceDataCount

    ' Skip the reserved data.
    lReservedDataCount = Read32(FILENUM_PICTURE)
    If lReservedDataCount <> 0 Then lFilePointer = lFilePointer + lReservedDataCount

    ' Find out if the data is compressed.
    .Compression = Read16(FILENUM_PICTURE)

    'Compression Type 0=Raw Data, RLE Compressed = 1
    'Compression Type Not Supported
    If .Compression > 1 Then Close FILENUM_PICTURE: Exit Sub

    ' Decode Image...
    ReDim .Pixels(0 To (4 * .Height * .Width) + 2) As Byte

    DecodePSD FILENUM_PICTURE
End With

Close FILENUM_PICTURE

'Copy this data into our custom image object (which was passed)
With pImage
    .BPP = 24
    .Width = ImageInfo.Width
    .Height = ImageInfo.Height

    Erase .Palette
    ReDim .Data(1 To ((.Width * .Height) * 3))

    lOffset1 = 1
    lOffset2 = 0

    For nIndex = 0 To ImageInfo.Height - 1
        For nIndex2 = 0 To ImageInfo.Width - 1
            .Data(lOffset1) = ImageInfo.Pixels((lOffset2 * 4))

            lOffset1 = lOffset1 + 1

            .Data(lOffset1) = ImageInfo.Pixels((lOffset2 * 4) + 1)

            lOffset1 = lOffset1 + 1

            .Data(lOffset1) = ImageInfo.Pixels((lOffset2 * 4) + 2)

            lOffset1 = lOffset1 + 1

            'Skip Alpha Pixel (which would be (lOffset2 *4) +1)
            lOffset2 = lOffset2 + 1
        Next nIndex2
    Next nIndex
End With

Erase ImageInfo.Pixels

Exit Sub

FileError:
ErrHandle
On Error GoTo 0
Close FILENUM_PICTURE

' Clear any information already in pImage
Dim pBlank As IMAGEFILE
pImage = pBlank

End Sub
Private Sub DecodePSD(iFileNumber As Integer)

On Error GoTo ErrorHandler

Dim lDefault(0 To 3) As Long
Dim lChannel(0 To 3) As Long
Dim lPixelCount As Long
Dim nIndex As Long, nIndex2 As Long
Dim lPointer As Long
Dim lCurrentChannel As Long
Dim lCount As Long
Dim lLength As Long
Dim bValue As Byte
Dim bFileContainer() As Byte

lDefault(0) = 0
lDefault(1) = 0
lDefault(2) = 0
lDefault(3) = 255

lChannel(0) = 2
lChannel(1) = 1
lChannel(2) = 0
lChannel(3) = 3

ReDim bFileContainer(0 To LOF(iFileNumber) - lFilePointer)

Get iFileNumber, lFilePointer, bFileContainer

lFilePointer = 0
lPixelCount = ImageInfo.Width * ImageInfo.Height

If ImageInfo.Compression Then
    lFilePointer = lFilePointer + ImageInfo.Height * ImageInfo.ChannelCount * 2

    For nIndex = 0 To 3
        lPointer = 0
        lCurrentChannel = lChannel(nIndex)
        If lCurrentChannel >= ImageInfo.ChannelCount Then
            For lPointer = 0 To lPixelCount - 1
                ImageInfo.Pixels((lPointer * 4) + lCurrentChannel) = lDefault(lCurrentChannel)
            Next lPointer
        Else
            lCount = 0

            Do Until (lCount >= lPixelCount)
                lLength = bFileContainer(lFilePointer)
                lFilePointer = lFilePointer + 1

                If lLength = 128 Then
                ElseIf lLength < 128 Then
                    lLength = lLength + 1
                    lCount = lCount + lLength

                    Do Until lLength = 0
                        ImageInfo.Pixels((lPointer * 4) + lCurrentChannel) = bFileContainer(lFilePointer)
                        lFilePointer = lFilePointer + 1
                        lPointer = lPointer + 1
                        lLength = lLength - 1
                    Loop
                ElseIf lLength > 128 Then
                    lLength = lLength Xor 255
                    lLength = lLength + 2
                    bValue = bFileContainer(lFilePointer)
                    lFilePointer = lFilePointer + 1
                    lCount = lCount + lLength

                    Do Until lLength = 0
                        ImageInfo.Pixels((lPointer * 4) + lCurrentChannel) = bValue
                        lPointer = lPointer + 1
                        lLength = lLength - 1
                    Loop
                End If
            Loop
        End If
    Next nIndex
Else
    For nIndex = 0 To 3
        lCurrentChannel = lChannel(nIndex)

        If lCurrentChannel > ImageInfo.ChannelCount Then
            For lPointer = 0 To lPixelCount - 1
                ImageInfo.Pixels((lPointer * 4) + lCurrentChannel) = lDefault(lCurrentChannel)
            Next lPointer
        Else
            For nIndex2 = 0 To lPixelCount - 1
                ImageInfo.Pixels((nIndex2 * 4) + lCurrentChannel) = bFileContainer(lFilePointer)
                lFilePointer = lFilePointer + 1
            Next nIndex2
        End If
    Next nIndex
End If

Erase bFileContainer

Exit Sub

ErrorHandler:
ErrHandle
On Error GoTo 0

End Sub
