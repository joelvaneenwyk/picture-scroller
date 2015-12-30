VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lWndHandle As Long
Dim iTrackNum As Integer
Dim bMusicType As Byte

Const MUSIC_WAVE = 1
Const MUSIC_CD = 2
Const MUSIC_MIDI = 3

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Property Let hwnd(ByVal lNewValue As Long)

lWndHandle = lNewValue

End Property
Public Sub PlayMusic()

' Purpose: Play background music, either a file or a
'   CD, depending on the user's selection.
'   Note: Error handling is in this sub, because if
'     something goes wrong, the user would like to know,
'     but in other places, it is unnecessary.

Dim lReturnVal As Long
Dim sMCIReturn As String
Dim sFileName As String

If tProgramOptions.bMusicType = 0 Then Exit Sub

If bMusicType <> MUSIC_WAVE Then
    ' Get the current status of the device.
    sMCIReturn = Space(10)
    lReturnVal = mciSendString("Status BackMusic mode", sMCIReturn, Len(sMCIReturn), 0)
    sMCIReturn = UCase(TrimNulls(sMCIReturn))

    ' If it is just paused, then resume playing.
    If sMCIReturn = "PAUSED" Then
        lReturnVal = mciSendString("Play BackMusic notify", "", 0, 0)
        Exit Sub
    End If
End If

Select Case tProgramOptions.bMusicType
    ' Play a music file
    Case MUSIC_CD
        ' Open the CD
        lReturnVal = mciSendString("Open CDAudio ALIAS BackMusic wait", "", 0, 0)
        If lReturnVal <> 0 Then GoTo MCI_ERROR

        ' Set the time format to allow us to jump to tracks
        lReturnVal = mciSendString("Set BackMusic time format tmsf wait", "", 0, 0)
        If lReturnVal <> 0 Then GoTo MCI_ERROR

        If tProgramOptions.iTrackNumber <> 0 Then
            ' Play the specified track.
            lReturnVal = mciSendString("Play BackMusic from " & tProgramOptions.iTrackNumber & " to " & IIf(tProgramOptions.iTrackNumber = iTrackNum, iTrackNum, tProgramOptions.iTrackNumber + 1) & " notify", "", 0, lWndHandle)
            If lReturnVal <> 0 Then GoTo MCI_ERROR
        Else
            ' Start playing from the beginning.
            lReturnVal = mciSendString("Seek BackMusic to start", "", 0, 0)
            If lReturnVal <> 0 Then GoTo MCI_ERROR
            lReturnVal = mciSendString("Play BackMusic notify", "", 0, lWndHandle)
            If lReturnVal <> 0 Then GoTo MCI_ERROR
        End If
    Case Else
        If tProgramOptions.sMusicFile = "" Then Exit Sub

        sFileName = tProgramOptions.sMusicFile

        ' Add quotes around the file name
        If Left$(sFileName, 1) <> Chr$(34) Then sFileName = Chr$(34) & sFileName
        If Right$(sFileName, 1) <> Chr$(34) Then sFileName = sFileName & Chr$(34)

        ' Attempt to open the file as a midi file.
        ' If it works then the file must be an midi file.
        lReturnVal = mciSendString("Open " & sFileName & " TYPE sequencer ALIAS BackMusic wait", "", 0, 0)

        If lReturnVal = 0 Then
            ' Since it worked, it must be a midi file,
            ' so go ahead and play it.
            bMusicType = MUSIC_MIDI

            ' Begin playing the file
            lReturnVal = mciSendString("Play BackMusic notify", "", 0, lWndHandle)
            If lReturnVal <> 0 Then GoTo MCI_ERROR
        Else
            ' It didn't work, so it must be a wave file.
            bMusicType = MUSIC_WAVE

            ' DirectSound doesn't want quotes around
            ' a file name.
            ' NOTE: You would think we could just save
            '   a copy of the file name without the
            '   quotes before we put them in in the first
            '   place, but the thing is, what if there
            '   were quotes originally.
            sFileName = Mid$(sFileName, 2, Len(sFileName) - 2)

            DirectSound.CreateStreamingBuffer sFileName
            DirectSound.PlaySound "STREAM", tProgramOptions.bLoopMusic
        End If
End Select

Exit Sub

MCI_ERROR:
    MsgBox GetMCIError(lReturnVal)
    MsgBox "An error occurred while trying play the music.", vbCritical, "Error..."
    lReturnVal = mciSendString("Close BackMusic", "", 0, 0)

End Sub
Public Sub PauseMusic()

mciSendString "Pause BackMusic", "", 0, 0

End Sub
Public Sub StopMusic()

If bMusicType <> MUSIC_WAVE Then
    mciSendString "Stop BackMusic", "", 0, 0
    mciSendString "Close BackMusic", "", 0, 0
Else: DirectSound.StopSound "STREAM"
End If

End Sub
Public Function GetTrackNumber() As Byte

Dim lReturnVal As Long
Dim sMCIReturn As String
Dim nIndex As Integer

Screen.MousePointer = 11

' If this doesn't work, someone needs to know.
iTrackNum = 0

' Close any previous CD work
lReturnVal = mciSendString("Close MusicCD", "", 0, 0)

' Attempt to open a CDAudio "thingy"
lReturnVal = mciSendString("Open CDAudio ALIAS MusicCD shareable", "", 0, 0)
If lReturnVal <> 0 Then GoTo MCI_ERROR

' Detect if a CD is present
sMCIReturn = Space(5)
lReturnVal = mciSendString("Status MusicCD media present", sMCIReturn, 25, 0)
If lReturnVal <> 0 Then GoTo MCI_ERROR
sMCIReturn = UCase(TrimNulls(sMCIReturn))

' If a CD is present...
If sMCIReturn = "TRUE" Then
    ' Get the number of tracks; this number includes
    ' data tracks, not just audio tracks
    sMCIReturn = Space(100)
    lReturnVal = mciSendString("Status MusicCD number of tracks", sMCIReturn, 100, 0)
    If lReturnVal <> 0 Then GoTo MCI_ERROR
    sMCIReturn = TrimNulls(sMCIReturn)

    ' Loop through the tracks and get all audio ones
    For nIndex = 1 To sMCIReturn
        sMCIReturn = Space(100)
        lReturnVal = mciSendString("Status MusicCD type track " & nIndex, sMCIReturn, 25, 0)
        If lReturnVal <> 0 Then GoTo MCI_ERROR

        ' If this IS an audio track then increment
        ' the track number variable.
        sMCIReturn = UCase(TrimNulls(sMCIReturn))
        If sMCIReturn = "AUDIO" Then
            iTrackNum = iTrackNum + 1
        End If
    Next nIndex
End If

' Close down our CDAudio "thingy"
lReturnVal = mciSendString("Close MusicCD", "", 0, 0)

If iTrackNum > 0 Then GetTrackNumber = True

Screen.MousePointer = 0

Exit Function

MCI_ERROR:
    Screen.MousePointer = 0
    MsgBox "An error occurred while trying to read CD, it may be that:" & vbCr & vbCr & "1. No sound card is installed." & vbCr & "2. Your system does not support CD audio.", vbCritical, "Error..."
    lReturnVal = mciSendString("Close MusicCD", "", 0, 0)

End Function
Public Property Get NumOfTracks() As Integer

NumOfTracks = iTrackNum

End Property
Public Sub CallBack()

' Purpose: This gets called by the message handling
'   procedure when the song has finished playing.

Dim lReturnVal As Long

' If looping is enabled...
If tProgramOptions.bLoopMusic = 0 Then
    Select Case tProgramOptions.bMusicType
        Case 1
            ' Go back to the beginning of the file
            lReturnVal = mciSendString("Seek BackMusic to start", "", 0, 0)
            lReturnVal = mciSendString("Play BackMusic notify", "", 0, lWndHandle)
        Case 2
            ' Play that track again
            lReturnVal = mciSendString("Play BackMusic from " & tProgramOptions.iTrackNumber & " to " & IIf(tProgramOptions.iTrackNumber = iTrackNum, iTrackNum, tProgramOptions.iTrackNumber + 1) & " notify", "", 0, lWndHandle)
    End Select
Else: StopMusic
End If

End Sub
Private Function GetMCIError(ByVal lError As Long) As String

Dim sBuffer As String

sBuffer = Space(255)

mciGetErrorString lError, sBuffer, 255

GetMCIError = TrimNulls(sBuffer)

End Function
