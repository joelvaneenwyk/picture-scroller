Attribute VB_Name = "modDeclarations"
Option Explicit

' This module contains all public declarations that do not
' fit in any other module or form.

Public Enum CREATE_METHODS
    CREATE_FROM_FILE = 1
    CREATE_FROM_RES = 2
    CREATE_FROM_HDC = 3
    CREATE_FROM_OLEPICTURE = 4
    CREATE_FROM_NONE = 5
End Enum

Public Const FILENUM_OPTIONS = 1
Public Const FILENUM_LIST = 2
Public Const FILENUM_PICTURE = 3
Public Const FILENUM_SOUND = 4

Public Const NUM_OF_TRANSITIONS = 50

Public Const LIST_HEADER = "--PICTURE SCROLLING SAVED LIST--"

Public Const HELP_FILE = "Picture Scroller.hlp"

Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1
Public Const HELP_QUIT = &H2

' Hold the size in pixels of the screen
Public iScreenWidth As Integer
Public iScreenHeight As Integer

' Says when an operation has been canceled or its done
Public bCancelOp As Byte

' Holds the current drawing position and picture
' dimensions while drawing on the screen
Public iPicLeft As Integer
Public iPicTop As Integer
Public iPicWidth As Integer
Public iPicHeight As Integer

' Defines a display mode.
Public Type DISPLAY_MODE
    iWidth As Integer
    iHeight As Integer
    bBPP As Byte
End Type

' These are the options that the user has set
Private Type PROGRAM_OPTIONS
    ' True if NOT the first time the program started
    bNotFirstStart As Byte

    ' Last directory used that contained picture files.
    sLastPicFolder As String
    ' Last directory used that contained music files.
    sLastMusicFolder As String
    ' Last directory used that contained saved lists.
    sLastSavedFolder As String

    ' Interval between pictures
    iInterval As Integer

    ' Transitions
    bTransitions() As Byte
    ' Randomly use transitions
    bRandomTransitions As Byte

    ' Scrolling direction (Forward, backward, random)
    bScrollDirection As Byte
    ' Size of pictures (Original, stretched, proportional)
    bPictureSize As Byte

    ' Type of background music (CD, file, none)
    bMusicType As Byte
    ' Path to background music file
    sMusicFile As String
    ' Loop music
    bLoopMusic As Byte
    ' Play a specific track or not
    bPlayTrack As Byte
    ' CD track number to play
    iTrackNumber As Integer

    ' Display mode to use when scrolling pictures
    bScreenSetting As Byte
    ' Display mode used by user
    tDisplayMode As DISPLAY_MODE
    ' Sound Effects
    bSoundEffects As Byte

    ' Color to draw picture information in
    lInfoColor As Long
    ' Custom colors the user has used before
    bCustomColors() As Byte
    ' Backdrop color for pictures
    lBackColor As Long

    ' Recent Saved lists
    sRecent1 As String
    sRecent2 As String

    ' Max length for paths.
    MAX_LEN As Integer
End Type

Public lSlot As Long
Public tPictureFiles As New Collection
Public tAddedFolders As New Collection

' Given us when we initialize HTML Help
Public lHelpCookie As Long

Public Type SELECTED_PIC
    lPictureIndex As Long
    tActualSize As RECT
    iValueX As Integer
    iValueY As Integer
End Type

Public tSelectedPics() As SELECTED_PIC

Public DirectX As New DirectX7
Public DirectDraw As New CDirectDraw
Public DirectSound As New CDirectSound
Public BackMusic As New CMusic

Public hEvent(2) As Long

' Surfaces that will always be needed.
Public dixuDirectDraw As DirectDraw7
Public dixuFrontSurface As DirectDrawSurface7
Public dixuBackBuffer As DirectDrawSurface7
Public dixuTransition As DirectDrawSurface7
Public dixuPicture As DirectDrawSurface7

Public tProgramOptions As PROGRAM_OPTIONS
