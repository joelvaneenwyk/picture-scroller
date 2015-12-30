VERSION 5.00
Begin VB.Form frmScroller 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Picture Scroller 2.0"
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "frmScroller.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3675
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   105
      Top             =   105
   End
   Begin VB.PictureBox picOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2625
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SCR_PICTURES = 1
Const SCR_OPTIONS = 2

Const SPI_SCREENSAVERRUNNING = 97

' Says what to print on the screen: pictures, options, etc.
Dim bCurrentScreen As Byte

Dim bHasTransitions             ' The user does want to use transitions
Dim bPause As Byte              ' Pause the transition and scrolling
Dim bUseTransition As Byte      ' Tell "MainLoop" to use a transition
Dim bDoingTransition As Byte    ' States that we're doing a transition
Dim bShowMouse As Byte          ' Should we show the mouse?
Dim bShowFPS As Byte            ' Show the frames/second
Dim bMinimized As Byte          ' The form is minimized.
Dim bUnload As Byte             ' Tells us to unload

' A random sequence of pictures to scroll through.
Dim lRandomNumbers() As Long

' Holds the last "display count" of the mouse.
Dim iMouseCount As Integer

' Position of the mouse (where to put the mouse)
Dim iMouseX As Integer
Dim iMouseY As Integer

' Class of the transitions for showing pictures
Dim Transitions As New CTransitions

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function VerifyScreenSavePwd Lib "Password.cpl" (ByVal hwnd As Long) As Long
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If bRunMode <> RM_SAVER_RUN Then
    ' If we're running in normal mode, then check all
    ' the usual keys.
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight
            bPause = True
            tmrScroll.Enabled = False

            TransitionDone True

            If KeyCode = vbKeyLeft Then
                LoadNextPicture "PREVIOUS", True
            ElseIf KeyCode = vbKeyRight Then
                LoadNextPicture "NEXT", True
            End If
        Case vbKeySpace
            ' Set to the opposite of the current value.
            bPause = Not bPause

            ' If we aren't in the middle of a transition,
            ' and we're not paused, then we should also
            ' re-enable the timer.
            If bUseTransition = False And bPause = False Then
                tmrScroll.Enabled = True
            Else: tmrScroll.Enabled = False
            End If
        Case vbKeyH
            bShowMouse = Not bShowMouse
        Case vbKeyF
            bShowFPS = Not bShowFPS
        Case vbKeyEscape
            bUnload = True
    End Select
Else: bUnload = True
End If

End Sub
Private Sub Form_Load()

' Find the original mouse count
iMouseCount = ShowCursor(False) + 1

' Set the mouse count to 0
Do While ShowCursor(False) >= -1
    DoEvents
Loop
Do While ShowCursor(True) <= -1
    DoEvents
Loop

' Hide the mouse.
ShowCursor False

Randomize

' If we're to show the pictures randomly then
' we need some numbers.
If tProgramOptions.bScrollDirection = 2 Then
    ' If the user cancel the operation, then just unload.
    If GenerateRandomNumbers = False Then Exit Sub
End If

lPrevWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf ScrollingProc)

' Initialize DirectDraw; exit if it doesn't work.
If DirectDraw.InitDirectDraw(Me) <> True Then Exit Sub

' Disable CTRL+ALT+DELETE if we're in ScreenSaver mode
' and this ISN'T Windows NT.
If bRunMode = RM_SAVER_RUN And bIsWinNT = False Then
    SystemParametersInfo SPI_SCREENSAVERRUNNING, True, 0, 0
End If

' Create the mouse cursor.
DirectDraw.CreateSurface "MOUSE", 0, 0, CREATE_FROM_RES, 1, "PICS"
DirectDraw.SetColorKey "MOUSE", RGB(255, 255, 255), RGB(255, 255, 255)

' Create the surface on which the pictures with be
' drawn before going to the back surface.
DirectDraw.CreateSurface "TRANSITION", iScreenWidth, iScreenHeight, CREATE_FROM_NONE, 0

bCurrentScreen = SCR_PICTURES

' Store whether to do transitions or not.
If UBound(tProgramOptions.bTransitions) = 0 Then
    bHasTransitions = False
Else: bHasTransitions = True
End If

' We can now start scrolling.
tmrScroll.Interval = tProgramOptions.iInterval

BackMusic.hwnd = Me.hwnd
BackMusic.PlayMusic

MainLoop

DirectDraw.KillDirectDraw

' We can now re-enable CTRL+ALT+DELETE, if it was
' previously disabled.
If bRunMode = RM_SAVER_RUN And bIsWinNT = False Then
    SystemParametersInfo SPI_SCREENSAVERRUNNING, False, 0, 0
ElseIf bRunMode = RM_SCROLL Then
    ' If we were just scrolling through the pictures,
    ' but not in Screen Saver mode, we can now just
    ' show frmMain, so that would mean that we're
    ' running in normal mode then.
    bRunMode = RM_NORMAL

    Load frmMain
    DoEvents
End If

End Sub
Private Sub MainLoop()

Dim iWidth As Integer
Dim iHeight As Integer

Dim iCounter As Integer
Dim fCurrentTime As Single
Dim fLastTime As Single
Dim FPS As Single

LoadNextPicture "NEXT", False

Do
    ' If the form has been minimized...
    If bMinimized = True Then
        ' Refresh the screen after us drawing on it
        DirectDraw.FlipToGDI

        ' Wait till the user brings us back
        Do While bMinimized = True
            DoEvents
        Loop

        DirectDraw.SetSurface "MOUSE", 0, 0, 0, 0, CREATE_FROM_RES, 1, "PICS"
        DirectDraw.RestoreSurfaces

        ' Clear the back buffer of any garbage
        DirectDraw.ClearRegion "BACKBUFFER"

        If bCurrentScreen = SCR_PICTURES Then
            ' Stop the last transition if this particular
            ' transition show redo after any interruption
            If Transitions.AlwaysRedo = True Then
                Transitions.StopTransition
            End If

            TransitionDone True

            ' Reshow the current picture
            LoadNextPicture "", False
        End If
    End If

    Select Case bCurrentScreen
        Case SCR_PICTURES
            If bPause = False Then
                If bUseTransition = True Then DoTransition
            End If

            ' Print the result of the transition or
            ' just the whole picture (if that's what
            ' we've done) onto the back buffer.
            DirectDraw.BltFast "BACKBUFFER", 0, 0, iScreenWidth, iScreenHeight, "TRANSITION", 0, 0, False
        Case SCR_OPTIONS
    End Select

    If bShowMouse = True Then
        If iMouseX > iScreenWidth - 15 Then
            iWidth = iScreenWidth - iMouseX
        Else: iWidth = 15
        End If

        If iMouseY > iScreenHeight - 25 Then
            iHeight = iScreenHeight - iMouseY
        Else: iHeight = 25
        End If

        DirectDraw.BltFast "BACKBUFFER", iMouseX, iMouseY, iWidth, iHeight, "MOUSE", 0, 0, True
    End If

    If bShowFPS = True Then
        If iCounter = 30 Then
            fCurrentTime = Timer
            If fLastTime <> 0 Then FPS = 30 / (fCurrentTime - fLastTime)
            fLastTime = fCurrentTime
            iCounter = 0
        End If

        DirectDraw.DrawText "BACKBUFFER", Format$(FPS, "###.00"), 5, 15, 100, 50, vbYellow

        iCounter = iCounter + 1
    End If

    If bUnload = True Then
        If bRunMode = RM_SAVER_RUN Then
            If VerifyExit = True Then Exit Do
        Else: Exit Do
        End If
    End If

    DirectDraw.Flip
    DirectDraw.ClearRegion "BACKBUFFER"

    DoEvents
Loop

Transitions.StopTransition

tmrScroll.Enabled = False

End Sub
Private Sub LoadNextPicture(ByVal Action As String, ByVal bJumpPicture As Byte)

Static lPictureIndex As Long
Dim lThisNumber As Long

' Do this so "MainLoop" does try to call "DoTransition"
bUseTransition = False

' Figure out what picture to show next
If Action = "NEXT" Then
    If tProgramOptions.bScrollDirection = 2 Then
        If lPictureIndex = UBound(lRandomNumbers) Then
            lPictureIndex = 1
        Else: lPictureIndex = lPictureIndex + 1
        End If

        lThisNumber = lRandomNumbers(lPictureIndex)
    Else
        If lPictureIndex = tPictureFiles.Count Then
            lPictureIndex = 1
        Else: lPictureIndex = lPictureIndex + 1
        End If

        lThisNumber = lPictureIndex
    End If
ElseIf Action = "PREVIOUS" Then
    If tProgramOptions.bScrollDirection = 2 Then
        If lPictureIndex = 1 Then
            lPictureIndex = UBound(lRandomNumbers)
        Else: lPictureIndex = lPictureIndex - 1
        End If

        lThisNumber = lRandomNumbers(lPictureIndex)
    Else
        If lPictureIndex = 1 Then
            lPictureIndex = tPictureFiles.Count
        Else: lPictureIndex = lPictureIndex - 1
        End If

        lThisNumber = lPictureIndex
    End If
End If

' Put the picture, whatever the type, into picOriginal
GetPicFromIndex picOriginal, lThisNumber

' Clear the temporary picture box for use.
picTemp.Picture = LoadPicture()

Select Case tProgramOptions.bPictureSize
    ' Stretch Proportionally
    Case 0
        iPicWidth = picOriginal.Width
        iPicHeight = picOriginal.Height

        If iPicWidth < iScreenWidth And iPicHeight < iScreenHeight Then
            Do Until iPicWidth = iScreenWidth Or iPicHeight = iScreenHeight
                iPicWidth = iPicWidth + 1
                iPicHeight = iPicHeight + 1
            Loop
        Else
            Do Until iPicWidth <= iScreenWidth And iPicHeight <= iScreenHeight
                iPicWidth = iPicWidth - 1
                iPicHeight = iPicHeight - 1
            Loop
        End If

        iPicLeft = (iScreenWidth - iPicWidth) \ 2
        iPicTop = (iScreenHeight - iPicHeight) \ 2

        picTemp.Width = iPicWidth
        picTemp.Height = iPicHeight

        StretchBlt picTemp.hdc, 0, 0, iPicWidth, iPicHeight, picOriginal.hdc, 0, 0, picOriginal.Width, picOriginal.Height, vbSrcCopy
        Set picTemp.Picture = picTemp.Image

    ' Original Size
    Case 1
        If picOriginal.Width > iScreenWidth Then
            iPicWidth = iScreenWidth
            iPicLeft = 0
        Else
            iPicWidth = picOriginal.Width
            iPicLeft = (iScreenWidth - iPicWidth) \ 2
        End If

        If picOriginal.Height > iScreenHeight Then
            iPicHeight = iScreenHeight
            iPicTop = 0
        Else
            iPicHeight = picOriginal.Height
            iPicTop = (iScreenHeight - iPicHeight) \ 2
        End If

        picTemp.Width = iPicWidth
        picTemp.Height = iPicHeight

        BitBlt picTemp.hdc, 0, 0, iPicWidth, iPicHeight, picOriginal.hdc, (picOriginal.Width - iPicWidth) \ 2, (picOriginal.Height - iPicHeight) \ 2, vbSrcCopy
        Set picTemp.Picture = picTemp.Image

    ' Stretch Full-Screen
    Case 2
        iPicWidth = iScreenWidth
        iPicHeight = iScreenHeight

        iPicLeft = 0
        iPicTop = 0

        picTemp.Width = iPicWidth
        picTemp.Height = iPicHeight

        StretchBlt picTemp.hdc, 0, 0, iPicWidth, iPicHeight, picOriginal.hdc, 0, 0, picOriginal.Width, picOriginal.Height, vbSrcCopy
        Set picTemp.Picture = picTemp.Image
End Select

DirectDraw.CreateSurface "CURRENTPICTURE", iPicWidth, iPicHeight, CREATE_FROM_HDC, picTemp

DirectDraw.ClearRegion "TRANSITION"

' Just print the picture onto the "TRANSITION" surface if
'   1. The user wants DOESN'T want to use transitions
'   2. We ARE told to just show the picture
'   3. The picture is SMALLER than 60x60
If bHasTransitions = False Or bJumpPicture = True Or iPicWidth < 60 Or iPicHeight < 60 Then
    ' Blt the picture onto the "TRANSITION" surface,
    ' which will then be blted by MainLoop from the
    ' back surface to the front surface
    DirectDraw.BltFast "TRANSITION", iPicLeft, iPicTop, iPicWidth, iPicHeight, "CURRENTPICTURE", 0, 0, False

    ' Clear everything.
    TransitionDone True

    ' If we were to just print the picture, then we
    ' shouldn't re-enable the timer.
    If bJumpPicture = False Then
        tmrScroll.Enabled = True
    End If
Else
    ' Tell "MainLoop" that they can now start calling
    ' the transition.
    bUseTransition = True
End If

End Sub
Private Sub DoTransition()

' Holds which transition is currently being used.
Static bCurrentTransition As Byte

' If we're not already in the middle of a
' transition, then we need to select which
' transition to use this time.
If bDoingTransition = False Then
    With tProgramOptions
        If .bRandomTransitions = 0 Then
            ' Select the next transition from the
            ' array.
            If bCurrentTransition = UBound(.bTransitions) Then bCurrentTransition = 0

            bCurrentTransition = bCurrentTransition + 1
        Else
            ' Generate a random transition from the
            ' possible choices.
            bCurrentTransition = Int(UBound(.bTransitions) * Rnd + 1)
        End If
    End With

    ' NOW we are it the middle of a transition
    bDoingTransition = True
End If

With Transitions
Select Case bCurrentTransition
        Case 1: .Blinds 0
        Case 2: .Blinds 1
        Case 3: .BoxIn
        Case 4: .BoxOut
        Case 5: .Smear
        Case 6: .SlideIn
        Case 7: .MoveLU 1
        Case 8: .MoveRD 1
        Case 9: .MoveLU 0
        Case 10: .MoveRD 0
    End Select
End With

End Sub
Public Sub TransitionDone(Optional ByVal bPartial As Byte)

If bPartial = False Then
    ' bPartial is TRUE if all we wanted to do is the
    ' last two lines.  If a transition was actually
    ' finished (bPartial is FALSE), then we can re-enable
    ' the timer.
    tmrScroll.Enabled = True
End If

Transitions.StopTransition

' This makes sure that "MainLoop" doesn't try to
' make us print the picture again; we need to wait
' until another picture is loaded.
bUseTransition = False
' This is so that next time "DoTransition" is called
' we won't use the same transition.
bDoingTransition = False

DirectDraw.RemoveSurface "CURRENTPICTURE"

'bUnload = True

End Sub
Public Sub Minimize(ByVal bState As Byte)

If bState = True Then
    ShowCursor True

    ' Store the last value of tmrScroll; when the
    ' user re-enabled us, we can set it back.
    frmScroller.tmrScroll.Tag = frmScroller.tmrScroll.Enabled
    frmScroller.tmrScroll.Enabled = False

    bMinimized = True
Else
    ShowCursor False

    ' If the user has previously minimized us, but has reenabled
    ' us, then tmrScroll.Tag will hold the previous state of the
    ' timer, which we can now return to that previous state.
    If tmrScroll.Tag <> "" Then
        tmrScroll.Enabled = tmrScroll.Tag
        tmrScroll.Tag = ""
    End If

    bMinimized = False
End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Simulate the pressing of keys with the mouse buttons.
'If Button = 1 Then
'    Form_KeyDown vbKeyLeft, 0
'ElseIf Button = 2 Then
'    Form_KeyDown vbKeyRight, 0
'End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static iLastX As Integer
Static iLastY As Integer

Const MOUSE_RESPOND As Byte = 5

If bRunMode = RM_SAVER_RUN Then
    If ((iLastX = 0) And (iLastY = 0)) Or ((Abs(iLastX - X) < MOUSE_RESPOND) And (Abs(iLastY - Y) < MOUSE_RESPOND)) Then
        iLastX = X
        iLastY = Y
    Else
        iLastX = 0
        iLastY = 0

        bUnload = True
    End If
End If

iMouseX = X
iMouseY = Y

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

BackMusic.StopMusic

SetWindowLong Me.hwnd, GWL_WNDPROC, lPrevWndProc

bPause = False
bUseTransition = False
bDoingTransition = False
bShowMouse = False
bUnload = False

' Set the mouse count to the original count
Do While ShowCursor(False) >= iMouseCount
Loop

Do While ShowCursor(True) <= iMouseCount
Loop

End Sub
Private Sub tmrScroll_Timer()

tmrScroll.Enabled = False
LoadNextPicture "NEXT", False

End Sub
Private Function VerifyExit() As Byte

' If this IS Windows NT, then we can just exit.
If bIsWinNT = True Then GoTo CanExit

' Show the mouse.
ShowCursor True

' Do this so that the confirm window will be shown.
DirectDraw.FlipToGDI

' If the user enters the right password, then we can exit.
If VerifyScreenSavePwd(Me.hwnd) = 1 Then GoTo CanExit

' If we're here, then it means that we CANNOT exit.
bUnload = False

' Now we can hide the mouse.
ShowCursor False

Exit Function

CanExit:
VerifyExit = True

End Function
Private Function GenerateRandomNumbers() As Byte

' Purpose: Generates a sequence of random pictures
'   to scroll through.

Dim nIndex As Long
Dim tNumbers As New Collection
Dim lRandom As Long

' If there are more than 500 pictures this could take a while.
If tPictureFiles.Count > 500 Then
    WaitProcess "Loading", True
End If

ReDim lRandomNumbers(1 To tPictureFiles.Count)

' Fill tNumbers a continuous sequence of numbers.
For nIndex = 1 To tPictureFiles.Count
    tNumbers.Add nIndex

    If bCancelOp = True Then GoTo Done

    DoEvents
Next nIndex

For nIndex = 1 To tPictureFiles.Count
    ' Randomly select a number from the collection of
    ' numbers.
    lRandom = Int(tNumbers.Count * Rnd + 1)

    ' Use that value from the collection as the number.
    lRandomNumbers(nIndex) = tNumbers(lRandom)

    ' Now that that number has been used, remove it.
    tNumbers.Remove lRandom

    If bCancelOp = True Then GoTo Done

    DoEvents
Next nIndex

' If we're here, then we accomplished our mission.
GenerateRandomNumbers = True

Done:
' We can now stop the wait process, whether we were
' successful or not.
EndWaitProcess

End Function
