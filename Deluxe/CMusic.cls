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

Dim dmLoader As DirectMusicLoader
Dim dmPerformance As DirectMusicPerformance
Dim dmSegment As DirectMusicSegment
Dim dmSegmentState As DirectMusicSegmentState

Dim lWndHandle As Long
Dim iTrackNum As Integer
Dim lPausePos As Long
Dim lStartPos As Long
Dim lPlayPos As Long

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Function InitMusic(ByVal hwnd As Long) As Byte

On Error GoTo ErrorHandler

Set dmLoader = DirectX.DirectMusicLoaderCreate
Set dmPerformance = DirectX.DirectMusicPerformanceCreate

dmPerformance.Init DirectSound.GetDirectSound, hwnd
dmPerformance.SetPort -1, 1

hEvent(2) = DirectX.CreateEvent(frmEvent)

dmPerformance.SetNotificationHandle hEvent(2)
dmPerformance.AddNotificationType DMUS_NOTIFY_ON_SEGMENT

InitMusic = True

Exit Function

ErrorHandler:
ErrHandle
On Error GoTo 0
InitMusic = False

End Function
Public Sub PlayMusic()

' Purpose: Play background music, either a file or a
'   CD, depending on the user's selection.
'   Note: Error handling is in this sub, because if
'     something goes wrong, the user would like to know,
'     but in other places, it is unnecessary.

On Error GoTo ErrorHandler

Dim lReturnVal As Long
Dim bMidiFile As Byte
Dim sMCIReturn As String

Select Case tProgramOptions.bMusicType
    ' Wave file or MIDI file
    Case 1
        If tProgramOptions.sMusicFile = "" Then Exit Sub

        If Not dmSegment Is Nothing Then
            ' If we've previously loaded a midi file,
            ' then see if its playing.
            If dmPerformance.IsPlaying(dmSegment, dmSegmentState) = False Then
                lPlayPos = lPausePos - lStartPos + lPlayPos + 1
                dmSegment.SetStartPoint lPlayPos
                Set dmSegmentState = dmPerformance.PlaySegment(dmSegment, 0, 0)
            End If

            ' Now we're satisfied (it's playing) so exit.
            Exit Sub
        ElseIf DirectSound.SoundExists("STREAM") = True Then
            If DirectSound.IsPlaying("STREAM") = False Then
                DirectSound.PlaySound "STREAM", IIf(tProgramOptions.bLoopMusic = 0, True, False)
            End If

            Exit Sub
        End If

        ' If we're still here, then it means that we
        ' haven't loaded a sound previously, so we need to.

        ' Attempt to load the file as a midi file.
        Set dmSegment = dmLoader.LoadSegment(tProgramOptions.sMusicFile)

        On Error Resume Next
        ' See if it REALLY loaded the file.  If an error
        ' occurred, then it really isn't a midi file.
        dmSegment.SetStandardMidiFile
        If Err.Number = 0 Then bMidiFile = True
        On Error GoTo 0

        If bMidiFile = True Then
            ' Set what note information to use.
            dmSegment.Download dmPerformance

            PlaySegment 0
        Else
            Set dmSegment = Nothing

            DirectSound.CreateStreamingBuffer tProgramOptions.sMusicFile
            DirectSound.PlaySound "STREAM", IIf(tProgramOptions.bLoopMusic = 0, True, False)
        End If
    ' CD Audio
    Case 2
        ' Get the current status of the device.
        sMCIReturn = Space(10)
        lReturnVal = mciSendString("Status BackMusic mode", sMCIReturn, Len(sMCIReturn), 0)
        sMCIReturn = UCase(TrimNulls(sMCIReturn))

        ' If it is just paused, then resume playing.
        If sMCIReturn = "PAUSED" Then
            lReturnVal = mciSendString("Play BackMusic notify", "", 0, 0)
        Else
            ' Open the CD
            lReturnVal = mciSendString("Open CDAudio ALIAS BackMusic wait", "", 0, 0)
            If lReturnVal <> 0 Then GoTo MCI_ERROR

            ' Set the time format to allow us to jump to tracks
            lReturnVal = mciSendString("Set BackMusic time format tmsf wait", "", 0, 0)
            If lReturnVal <> 0 Then GoTo MCI_ERROR

            ' Let us be notified when the track stops.
            lPrevWndProc = SetWindowLong(lWndHandle, GWL_WNDPROC, AddressOf MusicProc)

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
        End If
End Select

Exit Sub

MCI_ERROR:
mciSendString "Close BackMusic", "", 0, 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub PauseMusic()

' Purpose: Pauses the playing of the background music.

On Error GoTo ErrorHandler

Select Case tProgramOptions.bMusicType
    Case 1
        If Not dmSegment Is Nothing Then
            lStartPos = dmSegmentState.GetStartTime
            lPausePos = dmPerformance.GetMusicTime
            dmPerformance.Stop dmSegment, Nothing, 0, DMUS_SEGF_BEAT
        ElseIf DirectSound.SoundExists("STREAM") = True Then
            DirectSound.StopSound "STREAM"
        End If
    Case 2
        mciSendString "Pause BackMusic", "", 0, 0
End Select

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub StopMusic()

' Purpose: Sound playing the background music and removes
'   it from memory.

On Error GoTo ErrorHandler

Select Case tProgramOptions.bMusicType
    Case 1
        If Not dmSegment Is Nothing Then
            ' Pausing is the same as stopping except that here
            ' we remove the sound as well.
            PauseMusic
            Set dmSegmentState = Nothing
            Set dmSegment = Nothing
        ElseIf DirectSound.SoundExists("STREAM") = True Then
            DirectSound.StopSound "STREAM"
            DirectSound.RemoveSound "STREAM"
        End If
    Case 2
        mciSendString "Stop BackMusic", "", 0, 0
        mciSendString "Close BackMusic", "", 0, 0

        SetWindowLong lWndHandle, GWL_WNDPROC, lPrevWndProc
End Select

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Function GetTrackCount() As Byte

On Error GoTo ErrorHandler

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

If iTrackNum > 0 Then GetTrackCount = True

Screen.MousePointer = 0

Exit Function

MCI_ERROR:
    Screen.MousePointer = 0
    MsgBox "An error occurred while trying to read CD, it may be that:" & vbCr & vbCr & "1. No sound card is installed." & vbCr & "2. Your system does not support CD audio.", vbCritical, "Error..."
    lReturnVal = mciSendString("Close MusicCD", "", 0, 0)

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Property Get NumOfTracks() As Integer

NumOfTracks = iTrackNum

End Property
Public Sub CallBack()

' Purpose: This gets called by the message handling
'   procedure when the song has finished playing.

On Error GoTo ErrorHandler

' If looping is enabled...
If tProgramOptions.bLoopMusic = 0 Then
    ' Play that track again
    mciSendString "Play BackMusic from " & tProgramOptions.iTrackNumber & " to " & IIf(tProgramOptions.iTrackNumber = iTrackNum, iTrackNum, tProgramOptions.iTrackNumber + 1) & " notify", "", 0, lWndHandle
Else: StopMusic
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub MusicEvent(ByVal hEvent As Long)

' Purpose: Notification from DirectSound.

On Error GoTo ErrorHandler

Dim tMessage As DMUS_NOTIFICATION_PMSG

If dmPerformance.GetNotificationPMSG(tMessage) = True Then
    If tMessage.lNotificationOption = DMUS_NOTIFICATION_SEGALMOSTEND And tProgramOptions.bLoopMusic = 0 Then
        PlaySegment DMUS_SEGF_AFTERPREPARETIME
    End If
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function GetMCIError(ByVal lError As Long) As String

On Error GoTo ErrorHandler

Dim sBuffer As String

sBuffer = Space(255)

mciGetErrorString lError, sBuffer, 255

GetMCIError = TrimNulls(sBuffer)

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub PlaySegment(ByVal lFlags As Long)

On Error GoTo ErrorHandler

lStartPos = 0
lPlayPos = 0
lPausePos = 0

dmSegment.SetStartPoint 0

' Play the music.
Set dmSegmentState = dmPerformance.PlaySegment(dmSegment, lFlags, 0)

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Class_Terminate()

On Error GoTo ErrorHandler

If Not dmLoader Is Nothing Then
    StopMusic

    If hEvent(2) <> 0 Then DirectX.DestroyEvent hEvent(2)

    dmPerformance.CloseDown
    Set dmPerformance = Nothing
    Set dmLoader = Nothing
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
