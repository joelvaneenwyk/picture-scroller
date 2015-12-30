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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
Public Function InitDirectSound(ByVal hwnd As Long) As Byte

' Purpose: Initializes DirectSound objects.

On Error GoTo ErrorHandler

Dim dxEvent As DirectXEvent

' Initially think that this won't work.
bInitOK = False

lWndHandle = hwnd

Set dsDirectSound = DirectX.DirectSoundCreate("")
dsDirectSound.SetCooperativeLevel lWndHandle, DSSCL_PRIORITY

bInitOK = True
InitDirectSound = True

Exit Function

ErrorHandler:
ErrHandle
On Error GoTo 0
InitDirectSound = False

End Function
Public Function OpenSound(ByVal sKey As String, ByVal bGlobal As Byte, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant) As Byte

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

Set dsbSound = colSounds(sKey)
dsbSound.restore

dsbSound.Stop

dsbSound.Play IIf(bLoop = True, DSBPLAY_LOOPING, DSBPLAY_DEFAULT)

Set dsbSound = Nothing

Exit Sub

ErrorHandler:
ErrHandle
On Error GoTo 0

End Sub
Public Sub StopSound(ByVal sKey As String)

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

Set dsbSound = colSounds(sKey)
dsbSound.restore

dsbSound.Stop

Set dsbSound = Nothing

Exit Sub

ErrorHandler:
ErrHandle
On Error GoTo 0

End Sub
Public Sub RemoveSound(ByVal sKey As String)

colSounds.Remove sKey

End Sub
Public Property Get IsPlaying(ByVal sKey As String) As Byte

' Purpose: Returns whether or not a sound is playing.

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

Set dsbSound = colSounds(sKey)
dsbSound.restore

If dsbSound.GetStatus And DSBSTATUS_PLAYING Then
    IsPlaying = True
End If

Set dsbSound = Nothing

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Public Property Get SoundExists(ByVal sKey As String) As Byte

' Purpose: Returns whether or not sound exists.

On Error GoTo ErrorHandler

Dim dsbSound As DirectSoundBuffer

On Error Resume Next
Set dsbSound = colSounds(sKey)

' If we successfully loaded the sound then it is there.
If Not dsbSound Is Nothing Then SoundExists = True

Set dsbSound = Nothing
On Error GoTo 0

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Public Function GetDirectSound() As DirectSound

Set GetDirectSound = dsDirectSound

End Function
Private Sub Class_Terminate()

On Error GoTo ErrorHandler

While colSounds.Count > 0
    colSounds.Remove 1
Wend

Set dsDirectSound = Nothing

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
