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

Dim dsDirectSound As DirectSound
Dim colSounds As New Collection

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
Public Function InitDirectSound(ByVal hwnd As Long) As Byte

On Error GoTo ErrorHandler

lWndHandle = hwnd

DirectSoundCreate ByVal 0, dsDirectSound, Nothing
dsDirectSound.SetCooperativeLevel lWndHandle, DSSCL_NORMAL

InitDirectSound = True

Exit Function

ErrorHandler:
On Error GoTo 0
InitDirectSound = False

End Function
Public Function OpenSound(ByVal sKey As String, ByVal bCreateFrom As CREATE_METHODS, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant) As Byte

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
        ReDim bFileBytes(1 To FileLen(vSource))

        Open vSource For Binary As #1

        Get #1, , bFileBytes

        Close #1
    Case CREATE_FROM_RES
        sSoundData = LoadResData(vSource, vOtherInfo)
        ReDim bFileBytes(1 To LenB(sSoundData))
        bFileBytes = sSoundData
End Select

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
    .dwSize = Len(tBufferDesc)
    .dwFlags = DSBCAPS_CTRLDEFAULT
    .dwBufferBytes = lSize
    .lpwfxFormat = VarPtr(tWaveFormat)
End With

dsDirectSound.CreateSoundBuffer tBufferDesc, dsbSound, Nothing

dsbSound.Lock 0&, lSize, lPointer1, lLong1, lPointer2, lLong2, 0&

CopyMemory lPointer1, VarPtr(bFileBytes(lPosition + 4 + 4)), lLong1

If lLong2 <> 0 Then
    CopyMemory lPointer2, VarPtr(bFileBytes(lPosition + 4 + 4 + lLong1)), lLong2
End If

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

dsbSound.Play 0, 0, IIf(bLoop = True, DSBPLAY_LOOPING, 0)

Set dsbSound = Nothing

Exit Sub

ErrorHandler:
On Error GoTo 0

End Sub
Public Sub SetPan(ByVal sKey As String, ByVal lPan As Long)

Dim dsbSound As DirectSoundBuffer
Dim lFlags As Long

Set dsbSound = colSounds(sKey)
dsbSound.restore

dsbSound.SetPan lPan

Set dsbSound = Nothing

End Sub
Public Sub RemoveSound(ByVal sKey As String)

colSounds.Remove sKey

End Sub
Private Sub Class_Terminate()

While colSounds.Count > 0
    colSounds.Remove 1
Wend

Set dsDirectSound = Nothing

End Sub
