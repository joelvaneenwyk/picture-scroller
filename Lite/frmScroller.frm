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

Const FIRST_START = 0
Const NOT_FIRST = 1

' -----------------------------------------
' DirectDraw surfaces.
Dim dixuMouse As DirectDrawSurface7
Dim dixuButtons() As DirectDrawSurface7
' -----------------------------------------
' Says what to print on the screen: pictures, options, etc.
Dim bHasTransitions             ' The user does want to use transitions.
Dim bPause As Byte              ' Pause the transition and scrolling.
Dim bUseTransition As Byte      ' Tell "MainLoop" to use a transition.
Dim bDoingTransition As Byte    ' States that we're doing a transition.
Dim bShowMouse As Byte          ' Show the mouse.
Dim bShowFPS As Byte            ' Show the frames/second.
Dim bShowInfo As Byte           ' Show picture information.
Dim bShowHelp As Byte           ' Show the shortcut keys.
Dim bShowWallpaper As Byte      ' Show wallpaper form.
Dim bUnload As Byte             ' Tells us to unload.
Dim bPrint As Byte
' -----------------------------------------
' Buttons.
Const NUM_OF_BUTTONS = 4

Dim tButtons() As RECT
Dim bButtonOver As Byte

' -----------------------------------------
' Other miscellaneous variables.

' A random sequence of pictures to scroll through.
Dim lRandomNumbers() As Long

' Holds the last "display count" of the mouse.
Dim iMouseCount As Integer

' Position of the mouse (where to put the mouse)
Dim iMouseX As Integer
Dim iMouseY As Integer

Dim sPictureInfo As String

' Current picture shown and the last picture shown.
Public lPictureIndex As Long
Public lLastPicture As Long

' Class of the transitions for showing pictures
Dim Transitions As New CTransitions
Dim bCurrentTransition As Byte
Private Sub Form_Click()

On Error GoTo ErrorHandler

Select Case bButtonOver
    Case 1: Form_KeyDown vbKeyEscape, 0
    Case 2: Form_KeyDown vbKeyLeft, 0
    Case 3: Form_KeyDown vbKeyRight, 0
    Case 4: Form_KeyDown vbKeySpace, 0
End Select

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_DblClick()

On Error GoTo ErrorHandler

Form_Click

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

' If we're running in normal mode, then check all
' the usual keys.
Select Case KeyCode
    Case vbKeyLeft, vbKeyRight
        If bPrint = True Then Exit Sub

        bPause = True
        tmrScroll.Enabled = False

        TransitionDone True

        If tProgramOptions.bScrollDirection = 0 Then
            If KeyCode = vbKeyLeft Then
                LoadNextPicture "PREVIOUS", True
            ElseIf KeyCode = vbKeyRight Then
                LoadNextPicture "NEXT", True
            End If
        Else
            If KeyCode = vbKeyLeft Then
                LoadNextPicture "NEXT", True
            ElseIf KeyCode = vbKeyRight Then
                LoadNextPicture "PREVIOUS", True
            End If
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
    Case vbKeyH: bShowMouse = Not bShowMouse
    Case vbKeyF: bShowFPS = Not bShowFPS
    Case vbKeyF1: bShowHelp = Not bShowHelp
    Case vbKeyF2: bShowInfo = Not bShowInfo
    Case vbKeyF3: bShowWallpaper = True
    Case vbKeyEscape: bUnload = True
End Select

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim tFont As New StdFont

Dim nIndex As Byte
Dim iWidth As Integer
Dim iHeight As Integer

' Redim tButtons initially so that no errors occur
' until we can put some REAL info into tButtons.
ReDim tButtons(0)

' Set defaults if first start.
If tProgramOptions.bNotFirstStart = False Then SetDefaults

' Attempt to initialize DirectDraw.
' This also shows the form so that the user knows
' that things are happening.
If DirectDraw.InitDirectDraw(Me.hwnd, tProgramOptions.lBackColor) <> True Then
    Uninit

    MsgBox "An error occurred while trying to initialize DirectX.  Please be sure that you have DirectX 7.0 or greater installed on your computer and that no other program is currently using DirectX."
End If

SetMouseCount False

With tFont
    .bOld = True
    .SIZE = 16
    .Name = "Arial"
End With

dixuBackBuffer.SetForeColor tProgramOptions.lInfoColor
dixuBackBuffer.SetFontTransparency True

dixuFrontSurface.SetFont tFont
dixuFrontSurface.SetForeColor vbYellow

Set tFont = Nothing

DirectDraw.DrawText dixuFrontSurface, "Loading...", True

Randomize

' If we're to show the pictures randomly then
' we need some numbers.
If tProgramOptions.bScrollDirection = 2 Then
    ' If the user cancel the operation, then just unload.
    If GenerateRandomNumbers = False Then Exit Sub
End If

DirectSound.InitDirectSound Me.hwnd

' Load standard bitmaps
LoadStandard FIRST_START

' Calculate the area for the buttons.
ReDim tButtons(1 To NUM_OF_BUTTONS)

' Calculate the total area that will be occupied by
' the buttons.
iWidth = (DirectDraw.Width(dixuButtons(1)) * NUM_OF_BUTTONS) + ((NUM_OF_BUTTONS - 1) * 10)
iHeight = DirectDraw.Height(dixuButtons(1))

' Record the dimensions of the buttons.
For nIndex = 1 To NUM_OF_BUTTONS
    If nIndex = 1 Then
        With tButtons(nIndex)
            .Left = (iScreenWidth - iWidth) \ 2
            .Top = iScreenHeight - iHeight - 20
            ' These only hold the width and height, not the
            ' width/height plus the left/top.
            .Right = DirectDraw.Width(dixuButtons(1))
            .Bottom = iHeight
        End With
    Else
        With tButtons(nIndex)
            .Right = tButtons(1).Right
            .Bottom = tButtons(1).Bottom
            .Top = tButtons(1).Top
            .Left = tButtons(nIndex - 1).Left + 10 + .Right
        End With
    End If
Next nIndex

' Store whether to do transitions or not.
If UBound(tProgramOptions.bTransitions) = 0 Then
    bHasTransitions = False
Else: bHasTransitions = True
End If

' We can now start scrolling.
tmrScroll.Interval = tProgramOptions.iInterval

MainLoop

Uninit

' Save the user's options.
SaveUserOptions

' We can now re-enable CTRL+ALT+DELETE, if it was
' previously disabled.
If bRunMode = RM_SCROLL Then
    ' If we were just scrolling through the pictures,
    ' but not in Screen Saver mode, we can now just
    ' show frmMain, so that would mean that we're
    ' running in normal mode then.
    bRunMode = RM_NORMAL

    Load frmMain
    DoEvents
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub MainLoop()

On Error GoTo ErrorHandler

Dim tBigFont As New StdFont
Dim tSmallFont As New StdFont

Dim nIndex As Byte
Dim bButton As Byte

Dim iWidth As Integer
Dim iHeight As Integer

Dim iCounter As Integer
Dim fCurrentTime As Single
Dim fLastTime As Single
Dim FPS As Single

With tBigFont
    .bOld = True
    .SIZE = 12
    .Name = "Arial"
End With

With tSmallFont
    .bOld = False
    .SIZE = 11
    .Name = "Arial"
End With

lPictureIndex = 0

If tProgramOptions.bScrollDirection = 0 Then
    lLastPicture = 0
Else: lLastPicture = tPictureFiles.Count - 1
End If

LoadNextPicture "NEXT", False

Do
    If DirectDraw.TestState = False Then
        ShowCursor True
        tmrScroll.Tag = tmrScroll.Enabled
        tmrScroll.Enabled = False

        Do Until DirectDraw.TestState
            DoEvents
        Loop

        ShowCursor False

        ' Get the pictures back in the surfaces.
        LoadStandard NOT_FIRST

        ' Clear the surfaces of any garbage.
        DirectDraw.ClearSurface dixuBackBuffer
        DirectDraw.ClearSurface dixuTransition
        DirectDraw.ClearSurface dixuPicture

        ' Stop the last transition if this particular
        ' transition should redo after any interruption
        If Transitions.AlwaysRedo = True Then
            TransitionDone True
        Else: Set dixuPicture = Nothing
        End If

        ' Reshow the current picture
        LoadNextPicture "", False

        If Transitions.AlwaysRedo = True Then
            ' If the transition must start over then
            ' just put the whole picture on the screen.
            DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, iPicWidth, iPicHeight, dixuPicture, 0, 0, False
        Else
            ' Otherwise advance the transition;
            ' this only happens when the transition
            ' is one that reprint the whole picture
            ' every time.
            DoTransition
        End If

        If tmrScroll.Tag <> "" Then
            tmrScroll.Enabled = tmrScroll.Tag
            tmrScroll.Tag = ""
        End If
    End If

    ' If we're not paused then advance the transition.
    If bPause = False Then
        If bUseTransition = True Then DoTransition
    End If

    ' Print the result of the transition or
    ' just the whole picture (if that's what
    ' we've done) onto the back buffer.
    DirectDraw.BltFast dixuBackBuffer, 0, 0, iScreenWidth, iScreenHeight, dixuTransition, 0, 0, False

    bPrint = False

    If bShowMouse = True Then
        ' --------------------------------------
        ' Draw the buttons on the screen.
        For nIndex = 1 To NUM_OF_BUTTONS
            With tButtons(nIndex)
                bButton = nIndex

                If bButtonOver = nIndex Then bButton = bButton + NUM_OF_BUTTONS + 1

                If nIndex = NUM_OF_BUTTONS And bPause = True Then bButton = bButton + 1

                DirectDraw.BltFast dixuBackBuffer, .Left, .Top, .Right, .Bottom, dixuButtons(bButton), 0, 0, False
            End With
        Next nIndex
        ' --------------------------------------
        ' Make sure the mouse position has enough room
        ' for the width of the mouse picture.
        If iMouseX > iScreenWidth - 15 Then
            iWidth = iScreenWidth - iMouseX
        Else: iWidth = 15
        End If

        ' Make sure the mouse position has enough room
        ' for the height of the mouse picture.
        If iMouseY > iScreenHeight - 25 Then
            iHeight = iScreenHeight - iMouseY
        Else: iHeight = 25
        End If

        DirectDraw.BltFast dixuBackBuffer, iMouseX, iMouseY, iWidth, iHeight, dixuMouse, 0, 0, True
    End If

    If bShowFPS = True Then
        If iCounter = 30 Then
            fCurrentTime = Timer
            If fLastTime <> 0 Then FPS = 30 / (fCurrentTime - fLastTime)
            fLastTime = fCurrentTime
            iCounter = 0
        End If

        ' Don't put the FPS unto it is a value number.
        If FPS <> 0 Then
            dixuBackBuffer.DrawText 5, 15, Format$(FPS, "###.00"), False
        Else: dixuBackBuffer.DrawText 5, 15, "Calculating FPS...", False
        End If

        iCounter = iCounter + 1
    End If

    If bShowHelp = True Then
        ' Draw the title.
        dixuBackBuffer.SetFont tBigFont
        dixuBackBuffer.DrawText 10, 10, "Scrolling Shortcut Keys", False

        ' Draw the shortcut keys.
        dixuBackBuffer.SetFont tSmallFont
        dixuBackBuffer.DrawText 20, 40, "Space Bar - Pause/Resume Scrolling", False
        dixuBackBuffer.DrawText 20, 60, "H - Hide/Show Control Buttons", False
        dixuBackBuffer.DrawText 20, 80, "F1 - Show these shortcut keys.", False
        dixuBackBuffer.DrawText 20, 100, "F2 - Show the file information for each picture shown.", False
        dixuBackBuffer.DrawText 20, 120, "F3 - Make the currently shown picture your background (wallpaper).", False
    End If

    If bShowInfo = True Then
        dixuBackBuffer.SetFont tSmallFont

        dixuBackBuffer.DrawText 10, iScreenHeight - 30, sPictureInfo, False
    End If

    If bShowWallpaper = True Then
        ' Show the mouse.
        ShowCursor True

        ' Do this so that the form will be shown.
        DirectDraw.FlipToGDI

        ' Show the wallpaper form.
        frmWallpaper.Show

        Do Until frmWallpaper.bDone = True
            DoEvents
        Loop

        Unload frmWallpaper

        bShowWallpaper = False

        ' Now we can hide the mouse.
        ShowCursor False
    End If

    If bUnload = True Then Exit Do

    DirectDraw.Flip
    DirectDraw.ClearSurface dixuBackBuffer

    DoEvents
Loop

Transitions.StopTransition

tmrScroll.Enabled = False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub LoadNextPicture(ByVal Action As String, ByVal bJumpPicture As Byte)

On Error GoTo ErrorHandler

Dim iPosX As Integer
Dim iPosY As Integer
Dim iWidth As Integer
Dim iHeight As Integer

' Do this so "MainLoop" does not try to call "DoTransition"
bUseTransition = False

Beginning:

' Figure out what picture to show next
If Action = "NEXT" Then
    If tProgramOptions.bScrollDirection = 2 Then
        If lLastPicture = UBound(lRandomNumbers) Then
            lLastPicture = 1
        Else: lLastPicture = lLastPicture + 1
        End If

        lPictureIndex = lRandomNumbers(lLastPicture)
    Else
        If lLastPicture = tPictureFiles.Count Then
            lLastPicture = 1
        Else: lLastPicture = lLastPicture + 1
        End If

        lPictureIndex = lLastPicture
    End If
ElseIf Action = "PREVIOUS" Then
    If tProgramOptions.bScrollDirection = 2 Then
        If lLastPicture = 1 Then
            lLastPicture = UBound(lRandomNumbers)
        Else: lLastPicture = lLastPicture - 1
        End If

        lPictureIndex = lRandomNumbers(lLastPicture)
    Else
        If lLastPicture = 1 Then
            lLastPicture = tPictureFiles.Count
        ElseIf lLastPicture = 0 Then
            lLastPicture = 1
        Else: lLastPicture = lLastPicture - 1
        End If

        lPictureIndex = lLastPicture
    End If
Else
    If tProgramOptions.bScrollDirection = 2 Then
        lPictureIndex = lRandomNumbers(lLastPicture)
    Else: lPictureIndex = lLastPicture
    End If
End If

' Clear the temporary picture box for use.
picTemp.Picture = LoadPicture()

' Put the picture, whatever the type, into picOriginal;
' if any error occurs just go back to the top.
If GetPicFromIndex(picOriginal, lPictureIndex) = False Then GoTo Beginning

' Set sPictureInfo to a "readable" format of information.
sPictureInfo = tPictureFiles(lPictureIndex).FileName & " - " & picOriginal.Width & "x" & picOriginal.Height

Select Case tProgramOptions.bPictureSize
    ' Original Size
    Case 0
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

    ' Stretch Proportionally
    Case 1
        iPicWidth = iScreenWidth
        iPicHeight = iScreenHeight

        GetProportional picOriginal.Width, picOriginal.Height, iPicWidth, iPicHeight

        picTemp.Width = iPicWidth
        picTemp.Height = iPicHeight

        iPicLeft = (iScreenWidth - picTemp.Width) \ 2
        iPicTop = (iScreenHeight - picTemp.Height) \ 2

        SetStretchBltMode picTemp.hdc, STRETCH_DELETESCANS
        StretchBlt picTemp.hdc, 0, 0, iPicWidth, iPicHeight, picOriginal.hdc, 0, 0, picOriginal.Width, picOriginal.Height, vbSrcCopy
        Set picTemp.Picture = picTemp.Image

    ' Stretch Full-Screen
    Case 2
        iPicWidth = iScreenWidth
        iPicHeight = iScreenHeight

        iPicLeft = 0
        iPicTop = 0

        picTemp.Width = iPicWidth
        picTemp.Height = iPicHeight

        SetStretchBltMode picTemp.hdc, STRETCH_DELETESCANS
        StretchBlt picTemp.hdc, 0, 0, iPicWidth, iPicHeight, picOriginal.hdc, 0, 0, picOriginal.Width, picOriginal.Height, vbSrcCopy
        Set picTemp.Picture = picTemp.Image

    ' Tiled
    Case 3
        iPicWidth = iScreenWidth
        iPicHeight = iScreenHeight

        iPicLeft = 0
        iPicTop = 0

        picTemp.Width = iPicWidth
        picTemp.Height = iPicHeight

        If picOriginal.Width > iScreenWidth And picOriginal.Height > iScreenHeight Then
        Else
            Do
                If iScreenWidth - iPosX < picOriginal.Width Then
                    iWidth = iScreenWidth - iPosX
                Else: iWidth = picOriginal.Width
                End If

                If iScreenHeight - iPosY < picOriginal.Height Then
                    iHeight = iScreenWidth - iPosY
                Else: iHeight = picOriginal.Height
                End If

                BitBlt picTemp.hdc, iPosX, iPosY, picOriginal.Width, picOriginal.Height, picOriginal.hdc, 0, 0, vbSrcCopy

                If iPosX + picOriginal.Width < iScreenWidth Then
                    iPosX = iPosX + picOriginal.Width
                Else
                    iPosX = 0
                    iPosY = iPosY + picOriginal.Height
                End If

                If iPosY >= iScreenHeight Then Exit Do
            Loop

            DoEvents
        End If

        BitBlt picTemp.hdc, 0, 0, iPicWidth, iPicHeight, picOriginal.hdc, 0, 0, vbSrcCopy
        Set picTemp.Picture = picTemp.Image
End Select

DirectDraw.CreateSurface dixuPicture, iPicWidth, iPicHeight, CREATE_FROM_HDC, picTemp

' Clear the screen only if the picture size is original
' or proportional.
If tProgramOptions.bPictureSize = 0 Or tProgramOptions.bPictureSize = 1 Then
    DirectDraw.ClearSurface dixuTransition
End If

' Just print the picture onto the "TRANSITION" surface if
'   1. The user wants DOESN'T want to use transitions
'   2. We ARE told to just show the picture
'   3. The picture is SMALLER than 60x60
If bHasTransitions = False Or bJumpPicture = True Or iPicWidth < 60 Or iPicHeight < 60 Then
    bPrint = True

    ' Blt the picture onto the "TRANSITION" surface,
    ' which will then be blted by MainLoop from the
    ' back surface to the front surface
    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, iPicWidth, iPicHeight, dixuPicture, 0, 0, False

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

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub DoTransition()

On Error GoTo ErrorHandler

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
    Select Case tProgramOptions.bTransitions(bCurrentTransition)
        Case 1: .Blinds 0
        Case 2: .Blinds 1
        Case 3: .CircleIn
        Case 4: .Slide 1
        Case 5: .Slide 0
        Case 6: .Slide 3
        Case 7: .Slide 2
        Case 8: .Maze
        Case 9: .MoveLU
        Case 10: .Cubes
    End Select
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub TransitionDone(Optional ByVal bPartial As Byte)

On Error GoTo ErrorHandler

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

DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, iPicWidth, iPicHeight, dixuPicture, 0, 0, False

Set dixuPicture = Nothing

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub Minimize(ByVal bState As Byte)

On Error GoTo ErrorHandler

ShowCursor bState

If bState = True Then
    ' Store the last value of tmrScroll; when the
    ' user re-enabled us, we can set it back.
    frmScroller.tmrScroll.Tag = frmScroller.tmrScroll.Enabled
    frmScroller.tmrScroll.Enabled = False
Else
    ' If the user has previously minimized us, but has reenabled
    ' us, then tmrScroll.Tag will hold the previous state of the
    ' timer, which we can now return to that previous state.
    If tmrScroll.Tag <> "" Then
        tmrScroll.Enabled = tmrScroll.Tag
        tmrScroll.Tag = ""
    End If
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

' Simulate the pressing of keys with the mouse buttons.
If bButtonOver = 0 Then
    If Button = 1 Then
        Form_KeyDown vbKeyLeft, 0
    ElseIf Button = 2 Then
        Form_KeyDown vbKeyRight, 0
    End If
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Dim nIndex As Byte
Dim bOverButton As Byte

If UBound(tButtons) <> 0 Then
    For nIndex = 1 To NUM_OF_BUTTONS
        With tButtons(nIndex)
            If X > .Left And X < .Left + .Right And Y > .Top And Y < .Top + .Bottom Then
                bButtonOver = nIndex

                bOverButton = True
                Exit For
            End If
        End With
    Next nIndex

    If bOverButton = False Then bButtonOver = 0
End If

iMouseX = X
iMouseY = Y

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo ErrorHandler

' Reset variables.
bShowMouse = False
bShowFPS = False
bShowInfo = False
bShowHelp = False
bUseTransition = False
bDoingTransition = False
bPause = False
bUnload = False

bCurrentTransition = 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrScroll_Timer()

tmrScroll.Enabled = False

If tProgramOptions.bScrollDirection = 0 Then
    LoadNextPicture "NEXT", False
Else: LoadNextPicture "PREVIOUS", False
End If

End Sub
Private Function GenerateRandomNumbers() As Byte

' Purpose: Generates a sequence of random pictures
'   to scroll through.

On Error GoTo ErrorHandler

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

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub LoadStandard(ByVal bTime As Byte)

On Error GoTo ErrorHandler

Dim iButtonWidth As Long
Dim iButtonHeight As Long

Dim nIndex As Byte

' Calculate the dimensions of the buttons based on the
' idea that the buttons will look good at their original
' size at 1280x1024.  Separated so that we only need
' to use an integer.
iButtonWidth = iScreenWidth / 800 * 125
iButtonHeight = iScreenHeight / 600 * 49

If bTime = FIRST_START Then
    ' Create the mouse cursor.
    DirectDraw.CreateSurface dixuMouse, 0, 0, CREATE_FROM_RES, 1, "PICS"
    DirectDraw.SetColorKey dixuMouse, RGB(255, 255, 255)

    ' Create the surface on which the pictures with be
    ' drawn before going to the back surface.
    DirectDraw.CreateSurface dixuTransition, iScreenWidth, iScreenHeight, CREATE_FROM_NONE, 0

    ' Load the buttons, both regular and hovered.
    ReDim dixuButtons(1 To (NUM_OF_BUTTONS + 1) * 2)

    For nIndex = 1 To UBound(dixuButtons)
        DirectDraw.CreateSurface dixuButtons(nIndex), iButtonWidth, iButtonHeight, CREATE_FROM_RES, nIndex + 1, "PICS"
    Next nIndex
Else
    DirectDraw.RestoreSurfaces

    ' Reset the surfaces to their associated picture.
    DirectDraw.SetSurface dixuMouse, 0, 0, 0, 0, CREATE_FROM_RES, 1, "PICS"

    For nIndex = 1 To UBound(dixuButtons)
        DirectDraw.SetSurface dixuButtons(nIndex), 0, 0, 0, 0, CREATE_FROM_RES, nIndex + 1, "PICS"
    Next nIndex
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Uninit()

' Purpose: Uninitializes DirectDraw, DirectSound, and
'   background music.

On Error GoTo ErrorHandler

Dim nIndex As Byte

Set dixuMouse = Nothing

For nIndex = 1 To NUM_OF_BUTTONS * 2
    Set dixuButtons(nIndex) = Nothing
Next nIndex

Set dixuTransition = Nothing
Set dixuPicture = Nothing

DirectDraw.KillDirectDraw

Set DirectSound = Nothing

SetMouseCount True

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
