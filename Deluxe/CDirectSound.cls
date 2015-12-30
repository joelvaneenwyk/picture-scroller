VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirectSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lWndHandle As Long

' Allow people to check this to see if init worked.
Public bInitOK As Byte

Dim dsDirectSound As DirectSound
Dim colSounds As New Collection

Dim dsbStream As DirectSoundBuffer

Private Type FileHeader
    dwRiff As Long
    dwFileSize As Long
    dwWave As Long
    dwFormat As Long
    dwFormatLength As Long
End Type

Private Type FileFormat
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type

Private Type HeaderChunk
    dwType As Long
    dwLen As Long
End Type

Dim tBufferDesc As DSBUFFERDESC
Dim tHeader As FileHeader
Dim tWaveFormat As WAVEFORMATEX
Dim lDataLength As Long
Dim lNotificationSize As Long
Dim lLastBit As Long
Dim bSoundBuffer() As Byte
Dim bMustStream As Byte
Dim lAtTheEnd As Long
Dim sStreamFile As String

' Only used when playing a streaming sound file.
Dim bLoopSound As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32" (dest As Any, ByVal cbSize As Long)
Public Function InitDirectSound(ByVal hwnd As Long) As Byte

' Purpose: Initializes DirectSound objects.

On Error GoTo ErrorHandler

Dim dxEvent As DirectXEvent

' Initially think that this won't work.
bInitOK = False

lWndHandle = hwnd

Set dsDirectSound = DirectX.DirectSoundCreate("")
dsDirectSound.SetCooperativeLevel lWndHandle, DSSCL_PRIORITY

' Create two events.
hEvent(0) = DirectX.CreateEvent(frmEvent)
hEvent(1) = DirectX.CreateEvent(frmEvent)

bInitOK = True
InitDirectSound = True

Exit Function

ErrorHandler:
ErrHandle
On Error GoTo 0
InitDirectSound = False

End Function
Public Function OpenSound(ByVal sKey As String, ByVal bGlobal As Byte, ByVal bCreateFrom As CREATE_METHODS, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant) As Byte

' Purpose: Opens a sound file (or resource) for playing.

On Error GoTo ErrorHandler

Dim sSoundData As String
Dim dsbSound As DirectSoundBuffer
Dim tWaveFormat As WAVEFORMATEX
Dim lSize As Long
Dim lPosition As Long
Dim bFileBytes() As Byte
Dim tBufferDesc As DSBUFFERDESC
Dim lPointer1 As Long
Dim lPointer2 As Long
Dim lLong1 As Long
Dim lLong2 As Long

Select Case bCreateFrom
    Case CREATE_FROM_FILE
        ' If we're creating the sound buffer from a file,
        ' then make it streaming.
        CreateStreamingBuffer vSource
    Case CREATE_FROM_RES
        ' If we're loading the sound from the resourse
        ' file, then since it much be small, we can
        ' just load the whole thing into the buffer.
        sSoundData = LoadResData(vSource, vOtherInfo)
        ReDim bFileBytes(1 To LenB(sSoundData))
        bFileBytes = sSoundData

        lPosition = 1

        While Chr$(bFileBytes(lPosition)) + Chr$(bFileBytes(lPosition + 1)) + Chr$(bFileBytes(lPosition + 2)) <> "fmt"
            lPosition = lPosition + 1
        Wend

        CopyMemory VarPtr(tWaveFormat), VarPtr(bFileBytes(lPosition + 8)), Len(tWaveFormat)

        While Chr$(bFileBytes(lPosition)) + Chr$(bFileBytes(lPosition + 1)) + Chr$(bFileBytes(lPosition + 2)) + Chr$(bFileBytes(lPosition + 3)) <> "data"
            lPosition = lPosition + 1
        Wend

        CopyMemory VarPtr(lSize), VarPtr(bFileBytes(lPosition + 4)), Len(lSize)

        With tBufferDesc
            .lBufferBytes = lSize
            If bGlobal = True Then
                ' This will allow the sound to continue
                ' playing even after we lose focus.
                .lFlags = DSBCAPS_GLOBALFOCUS
            End If
        End With

        Set dsbSound = dsDirectSound.CreateSoundBuffer(tBufferDesc, tWaveFormat)

        dsbSound.WriteBuffer 0, lSize, bFileBytes(lPosition + 4 + 4), DSBLOCK_DEFAULT
End Select

colSounds.Add dsbSound, sKey
Set dsbSound = Nothing

OpenSound = True

Exit Function

ErrorHandler:
On Error GoTo 0
OpenSound = False

End Function
Public Sub PlaySound(ByVal sKey As String, ByVal bLoop As Byte)

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

If sKey <> "STREAM" Then
    Set dsbSound = colSounds(sKey)
    dsbSound.restore

    dsbSound.Stop

    dsbSound.Play IIf(bLoop = True, DSBPLAY_LOOPING, DSBPLAY_DEFAULT)

    Set dsbSound = Nothing
Else
    dsbStream.SetCurrentPosition 0

    lAtTheEnd = 0
    bLoopSound = IIf(tProgramOptions.bLoopMusic = 0, True, False)

    dsbStream.Play IIf(bMustStream = True, DSBPLAY_LOOPING, DSBPLAY_DEFAULT)
End If

Exit Sub

ErrorHandler:
ErrHandle
On Error GoTo 0

End Sub
Public Sub StopSound(ByVal sKey As String)

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

If sKey <> "STREAM" Then
    Set dsbSound = colSounds(sKey)
    dsbSound.restore

    dsbSound.Stop

    Set dsbSound = Nothing
Else: dsbStream.Stop
End If

Exit Sub

ErrorHandler:
ErrHandle
On Error GoTo 0

End Sub
Public Sub RemoveSound(ByVal sKey As String)

If sKey = "STREAM" Then
    Set dsbStream = Nothing
Else: colSounds.Remove sKey
End If

End Sub
Public Function CreateStreamingBuffer(ByVal sFileName As String) As Boolean

' Purpose: Setup a streaming bSoundBuffer.  Only one file
'   can be opened as a streaming bSoundBuffer at any time.

On Error GoTo ErrorHandler

Dim tCreateEvent(1) As DSBPOSITIONNOTIFY

' Close the file in case it is open.
Close FILENUM_SOUND

' Remove any previously created streaming bSoundBuffer
Set dsbStream = Nothing

sStreamFile = sFileName

' Get the file's tWaveFormat.
tWaveFormat = FillFormat()

If tWaveFormat.nFormatTag <> WAVE_FORMAT_PCM Then
    Close FILENUM_SOUND
    Exit Function
End If

With tBufferDesc
    lNotificationSize = (tWaveFormat.lSamplesPerSec * 2) \ 2
    lNotificationSize = lNotificationSize * 4

    .lBufferBytes = lNotificationSize * 4
    lLastBit = (lDataLength \ .lBufferBytes) * .lBufferBytes
    .lFlags = DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_CTRLPOSITIONNOTIFY
End With

Set dsbStream = dsDirectSound.CreateSoundBuffer(tBufferDesc, tWaveFormat)

If lDataLength >= tBufferDesc.lBufferBytes Then
    tCreateEvent(0).lOffset = (tBufferDesc.lBufferBytes) \ 2
    tCreateEvent(0).hEventNotify = hEvent(0)

    tCreateEvent(1).lOffset = (tBufferDesc.lBufferBytes - 1)
    tCreateEvent(1).hEventNotify = hEvent(1)

    ' Set the position notifications.
    dsbStream.SetNotificationPositions 2, tCreateEvent()
Else
    tCreateEvent(0).lOffset = lDataLength - 1
    tCreateEvent(0).hEventNotify = hEvent(0)

    ' Set the position notifications.
    dsbStream.SetNotificationPositions 1, tCreateEvent()
End If

' Resize the wave data bSoundBuffer to the size of
' the direct sound bSoundBuffer.
If lDataLength < tBufferDesc.lBufferBytes Then
    ' If the size of the file is actually quite small
    ' we don't need to stream the buffer.
    ReDim bSoundBuffer(lDataLength - 1)
    bMustStream = False
Else
    ' Otherwise, we need to keep putting bits in.
    ReDim bSoundBuffer(tBufferDesc.lBufferBytes - 1)
    bMustStream = True
End If

' Get our first chunk of sound data.
Get FILENUM_SOUND, , bSoundBuffer

dsbStream.WriteBuffer 0, UBound(bSoundBuffer), bSoundBuffer(0), DSBLOCK_DEFAULT

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub DirectSoundEvent(ByVal eventid As Long)

' Purpose: This gets called every time DirectSound
'   needs more sound data.

On Error GoTo ErrorHandler

Dim lStartSilence As Long
Dim lLenSilence As Long
Dim lPosition As Long

If bMustStream = True Then
    Select Case eventid
        Case hEvent(0)
            ' If this is the last section of data...
            If Loc(FILENUM_SOUND) > lLastBit Then
                lAtTheEnd = lAtTheEnd + 1

                ' Read the last bit of the sound file.
                Get FILENUM_SOUND, , bSoundBuffer
                lStartSilence = lDataLength - lLastBit
                lLenSilence = tBufferDesc.lBufferBytes - lStartSilence
                RtlZeroMemory bSoundBuffer(lStartSilence), lLenSilence
            Else
                ' Read some more of the sound file.
                Get FILENUM_SOUND, , bSoundBuffer
            End If

            lPosition = ((UBound(bSoundBuffer) + 1) \ 2)
            lPosition = lPosition + lPosition Mod 2
            dsbStream.WriteBuffer 0, lPosition, bSoundBuffer(0), DSBLOCK_DEFAULT
        Case hEvent(1)
            lPosition = ((UBound(bSoundBuffer) + 1) \ 2)
            dsbStream.WriteBuffer lPosition, lPosition, bSoundBuffer(0), DSBLOCK_DEFAULT

            If lAtTheEnd = 2 Then
                dsbStream.Stop

                ' If we were told originally to loop the
                ' entire sound, then now's the time to do it.
                If bLoopSound = True Then
                    CreateStreamingBuffer sStreamFile
                    PlaySound "STREAM", True
                End If
            End If
    End Select
ElseIf bLoopSound = True Then
    CreateStreamingBuffer sStreamFile
    PlaySound "STREAM", True
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Property Get IsPlaying(ByVal sKey As String) As Byte

' Purpose: Returns whether or not a sound is playing.

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

If sKey = "STREAM" Then
    If dsbStream Is Nothing Then
        IsPlaying = False
    ElseIf dsbStream.GetStatus And DSBSTATUS_PLAYING Then
        IsPlaying = True
    End If
Else
    Set dsbSound = colSounds(sKey)
    dsbSound.restore

    If dsbSound.GetStatus And DSBSTATUS_PLAYING Then
        IsPlaying = True
    End If

    Set dsbSound = Nothing
End If

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Public Property Get SoundExists(ByVal sKey As String) As Byte

' Purpose: Returns whether or not sound exists.

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

If sKey = "STREAM" Then
    If Not dsbStream Is Nothing Then SoundExists = True
Else
    On Error Resume Next
    Set dsbSound = colSounds(sKey)

    ' If we successfully loaded the sound then it is there.
    If Not dsbSound Is Nothing Then SoundExists = True

    Set dsbSound = Nothing
    On Error GoTo 0
End If

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Public Function GetDirectSound() As DirectSound

Set GetDirectSound = dsDirectSound

End Function
Private Function FillFormat() As WAVEFORMATEX

' Purpose: Reads the tHeader info from a wave file
'   and returns the a filled WAVEFORMATEX structure.

On Error GoTo ErrorHandler

Dim tFormat As FileFormat
Dim tChunk As HeaderChunk
Dim bTemp As Byte
Dim nIndex As Long

Open sStreamFile For Binary Access Read As FILENUM_SOUND

' Get the tHeader information.
Get FILENUM_SOUND, , tHeader

' Make sure the file meets certain qualifications.
If tHeader.dwRiff <> &H46464952 Or _
    tHeader.dwWave <> &H45564157 Or _
    tHeader.dwFormatLength < 16 Then
        Exit Function
End If

' Get the wave format information.
Get FILENUM_SOUND, , tFormat

' Bypass the extra format bytes.
For nIndex = 1 To tHeader.dwFormatLength - 16
    Get FILENUM_SOUND, , bTemp
Next nIndex

' Read in a chunk of data.
Get FILENUM_SOUND, , tChunk

Do While tChunk.dwType <> &H61746164
    For nIndex = 1 To tChunk.dwLen
        Get FILENUM_SOUND, , bTemp
    Next nIndex

    Get FILENUM_SOUND, , tChunk
Loop

lDataLength = tChunk.dwLen

' Fill in the information we know.
With FillFormat
    .lAvgBytesPerSec = tFormat.nAvgBytesPerSec
    .lExtra = 0
    .lSamplesPerSec = tFormat.nSamplesPerSec
    .nBitsPerSample = tFormat.wBitsPerSample
    .nBlockAlign = tFormat.nBlockAlign
    .nChannels = tFormat.nChannels
    .nFormatTag = tFormat.wFormatTag
End With

' The file is left open to keep the file read position
' at the start of the wave file data.

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub Class_Terminate()

On Error GoTo ErrorHandler

While colSounds.Count > 0
    colSounds.Remove 1
Wend

If hEvent(0) <> 0 Then DirectX.DestroyEvent hEvent(0)
If hEvent(1) <> 0 Then DirectX.DestroyEvent hEvent(1)

Set dsbStream = Nothing
Set dsDirectSound = Nothing

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
