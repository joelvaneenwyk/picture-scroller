Attribute VB_Name = "modMusic"
Option Explicit

Dim lWndHandle As Long
Dim sMusicFile As String
Dim iTrackNum As Integer
Dim bLoopMusic As Byte

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Property Let hwnd(ByVal lNewValue As Long)

lWndHandle = lNewValue

End Property
Public Property Let LoopMusic(ByVal bNewValue As Byte)

bLoopMusic = bNewValue

End Property
Public Property Let FileName(ByVal sNewValue As String)

sMusicFile = sNewValue

' Add quotes around the file name
If Left(sMusicFile, 1) <> Chr(34) Then sMusicFile = Chr(34) & sMusicFile
If Right(sMusicFile, 1) <> Chr(34) Then sMusicFile = sMusicFile & Chr(34)

End Property
Public Sub PlayMusic()

Dim lReturnVal As Long

Select Case tProgramOptions.bMusicType
    ' Play a music file
    Case 1
        If sMusicFile = "" Then Exit Sub

        ' Open the music file
        lReturnVal = mciSendString("Open " & sMusicFile & " ALIAS BackMusic wait", "", 0, 0)
        If lReturnVal <> 0 Then GoTo MCI_ERROR

        ' Begin playing the file
        lReturnVal = mciSendString(AddCommands("Play BackMusic"), "", 0, lWndHandle)
        If lReturnVal <> 0 Then GoTo MCI_ERROR
    ' Play a CD
    Case 2
        ' Open the CD
        lReturnVal = mciSendString("Open CDAudio ALIAS BackMusic wait", "", 0, 0)
        If lReturnVal <> 0 Then GoTo MCI_ERROR

        ' Set the time format to allow us to jump to tracks
        lReturnVal = mciSendString("Set BackMusic time format tmsf wait", "", 0, 0)
        If lReturnVal <> 0 Then GoTo MCI_ERROR

        If tProgramOptions.iTrackNumber <> 0 Then
            ' Start playing from a specific track, and
            ' keep looping it.
            lReturnVal = mciSendString(AddCommands("Play BackMusic from " & tProgramOptions.iTrackNumber), "", 0, lWndHandle)
            If lReturnVal <> 0 Then GoTo MCI_ERROR
        Else
            ' Start playing from the beginning.
            lReturnVal = mciSendString(AddCommands("Play BackMusic"), "", 0, lWndHandle)
            If lReturnVal <> 0 Then GoTo MCI_ERROR
        End If
End Select

bPlaying = True

Exit Sub

MCI_ERROR:
    MsgBox "An error occurred while trying play the music.", vbCritical, "Error..."
    lReturnVal = mciSendString("Close BackMusic", "", 0, 0)

End Sub
Public Sub PauseMusic()

mciSendString "Pause BackMusic", "", 0, 0

End Sub
Public Sub StopMusic()

mciSendString "Stop BackMusic", "", 0, 0
mciSendString "Close BackMusic", "", 0, 0

End Sub
Public Function ReadCD() As Byte

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
sMCIReturn = Space(25)
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

If iTrackNum > 0 Then
    ReadCD = True
Else: ReadCD = False
End If

Screen.MousePointer = 0

Exit Function

MCI_ERROR:
    Screen.MousePointer = 0
    MsgBox "An error occurred while trying to read CD, it may be that:" & vbCr & vbCr & "1. No sound card is installed" & vbCr & "1. Your system does not support CD audio", vbCritical, "Error..."
    lReturnVal = mciSendString("Close MusicCD", "", 0, 0)

End Function
Public Property Get NumOfTracks() As Integer

NumOfTracks = iTrackNum

End Property
Private Function AddCommands(ByVal sCommandString As String) As String

' Purpose: Appends " notify" to the end of sCommandString
'   if we are to loop through the music

If bLoopMusic = True Then
    AddCommands = sCommandString & " notify"
Else: AddCommands = sCommandString
End If

End Function
Public Sub CallBack()

End Sub
