VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirectDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Handle of the window owning us.
Dim lWndHandle As Long

' Holds a list of supported display modes.
Dim tDisplayModes() As DISPLAY_MODE

Dim RedShiftLeft As Long
Dim RedShiftRight As Long
Dim GreenShiftLeft As Long
Dim GreenShiftRight As Long
Dim BlueShiftLeft As Long
Dim BlueShiftRight As Long

Dim lBackColor As Long

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Public Function InitDirectDraw(ByVal hwnd As Long, ByVal lClearColor As Long) As Byte

' Purpose: Initializes DirectDraw.  This includes creating
'   a front surface and getting the back surface.

On Error GoTo ErrorHandler

Dim tNewDisplayMode As DISPLAY_MODE
Dim tSurface As DDSURFACEDESC2
Dim tBlankSurface As DDSURFACEDESC2
Dim tSurfaceCaps As DDSCAPS2

lWndHandle = hwnd

AlwaysOnTop lWndHandle

Set dixuDirectDraw = Nothing

Set dixuDirectDraw = DirectX.DirectDrawCreate("")

dixuDirectDraw.SetCooperativeLevel lWndHandle, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT

GetDisplayModes lWndHandle

Select Case tProgramOptions.bScreenSetting
    Case 0
        ' If the user just wanted to keep the current
        ' resolution, then just set the variable so
        ' that we know.
        With tNewDisplayMode
            dixuDirectDraw.GetDisplayMode tSurface

            .iWidth = tSurface.lWidth
            .iHeight = tSurface.lHeight
        End With
    Case 1
        ' Just select the first display mode; it will be
        ' the fastest.
        tNewDisplayMode = tDisplayModes(1)
    Case 2
        ' Just select the last display mode; it will have
        ' the best quality.
        tNewDisplayMode = tDisplayModes(UBound(tDisplayModes))
    Case 3
        tNewDisplayMode = tProgramOptions.tDisplayMode
End Select

With tNewDisplayMode
    ' Set the screen resolution, if the user has choosen
    ' to do so.
    If tProgramOptions.bScreenSetting <> 0 Then
        dixuDirectDraw.SetDisplayMode .iWidth, .iHeight, .bBPP, 0, DDSDM_DEFAULT
    End If

    iScreenWidth = .iWidth
    iScreenHeight = .iHeight
End With

tSurface = tBlankSurface

With tSurface
    .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    .lBackBufferCount = 1
End With

' Create the front surface
Set dixuFrontSurface = dixuDirectDraw.CreateSurface(tSurface)

dixuFrontSurface.restore

' Create the backbuffer.
tSurfaceCaps.lCaps = DDSCAPS_BACKBUFFER
Set dixuBackBuffer = dixuFrontSurface.GetAttachedSurface(tSurfaceCaps)

dixuBackBuffer.restore

ClearSurface dixuBackBuffer
ClearSurface dixuFrontSurface

dixuFrontSurface.GetSurfaceDesc tSurface

' Set the background color.
If tSurface.ddpfPixelFormat.lRGBBitCount > 8 Then
    MaskToShiftValues tSurface.ddpfPixelFormat.lRBitMask, RedShiftRight, RedShiftLeft
    MaskToShiftValues tSurface.ddpfPixelFormat.lGBitMask, GreenShiftRight, GreenShiftLeft
    MaskToShiftValues tSurface.ddpfPixelFormat.lBBitMask, BlueShiftRight, BlueShiftLeft

    lBackColor = TranslateColor(lClearColor)
Else: lBackColor = lClearColor
End If

InitDirectDraw = True

Exit Function

ErrorHandler:
ErrHandle
On Error GoTo 0
InitDirectDraw = False

End Function
Public Sub GetDisplayModes(ByVal hwnd As Long)

' Purpose: Retrieves a list of all the supported display
'   modes on this computer.  It does this by telling
'   DirectDraw to call us back for each display mode, by
'   calling ModesCallback.

On Error GoTo ErrorHandler

Dim bNotInit As Byte
Dim dixuModesEnum As DirectDrawEnumModes
Dim tDisplayMode As DDSURFACEDESC2
Dim nIndex As Integer

If dixuDirectDraw Is Nothing Then
    ' If we haven't previously initialized DirectDraw,
    ' then we want to un-initialized what we do here.
    bNotInit = True

    ' Initialize a temporary "thingy" for checking display modes
    Set dixuDirectDraw = DirectX.DirectDrawCreate("")
    dixuDirectDraw.SetCooperativeLevel hwnd, DDSCL_NORMAL
End If

Set dixuModesEnum = dixuDirectDraw.GetDisplayModesEnum(0, tDisplayMode)

ReDim tDisplayModes(1 To dixuModesEnum.GetCount)

For nIndex = 1 To dixuModesEnum.GetCount
    dixuModesEnum.GetItem nIndex, tDisplayMode

    With tDisplayModes(nIndex)
        .iWidth = tDisplayMode.lWidth
        .iHeight = tDisplayMode.lHeight
        .bBPP = tDisplayMode.ddpfPixelFormat.lRGBBitCount
    End With
Next nIndex

If bNotInit = True Then Set dixuDirectDraw = Nothing

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Property Get ModeWidth(ByVal iIndex As Integer) As Integer

' Purpose: Returns the width of a specific display mode.

On Error GoTo ErrorHandler

ModeWidth = tDisplayModes(iIndex).iWidth

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Property Get ModeCount() As Integer

' Purpose: Returns the number of supported display modes.

On Error GoTo ErrorHandler

ModeCount = UBound(tDisplayModes)

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Property Get ModeHeight(ByVal iIndex As Integer) As Integer

' Purpose: Returns the height of a specific display mode.

On Error GoTo ErrorHandler

ModeHeight = tDisplayModes(iIndex).iHeight

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Property Get ModeBPP(ByVal iIndex As Integer) As Byte

' Purpose: Returns the number of colors of the specified
'   display mode.

On Error GoTo ErrorHandler

ModeBPP = tDisplayModes(iIndex).bBPP

Exit Property

ErrorHandler:
ErrHandle
Resume Next

End Property
Public Sub CreateSurface(dixuNewSurface As DirectDrawSurface7, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal bCreateFrom As CREATE_METHODS, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant)

' Purpose: Creates a surface of the specified dimensions.
'   If bCreateFrom is not CREATE_FROM_NONE, then we also
'   load the picture from the source into objPicture.
'   If no dimensions are specified, we can take the sizes
'   from objPicture.  Again, if bCreateFrom is not
'   CREATE_FROM_NONE, we pass what we know (objPicture)
'   to SetPicture so that they can put the picture on
'   the surface.

On Error GoTo ErrorHandler

Dim objPicture As New StdPicture

' Surface description
Dim tSurface As DDSURFACEDESC2

If bCreateFrom = CREATE_FROM_NONE Then
    ' If we aren't supposed to put a picture on the
    ' surface and we aren't even given the dimensions,
    ' then we cannot do anything, so exit.
    If iWidth = 0 Or iHeight = 0 Then Exit Sub
Else
    ' Load picture from specified location
    If bCreateFrom = CREATE_FROM_HDC Then
        Set objPicture = vSource.Picture
    ElseIf bCreateFrom = CREATE_FROM_RES Then
        Set objPicture = PictureFromBits(LoadResData(vSource, vOtherInfo))
    Else: Exit Sub
    End If

    ' Retrieve the actual dimensions of the picture
    ' if none are given.  This assumes that the picture
    ' is not bigger than the screen.  If the picture is
    ' bigger than the screen, an error will occur.
    If iWidth = 0 Then iWidth = HimetricToPixel(objPicture.Width, Screen.TwipsPerPixelX)
    If iHeight = 0 Then iHeight = HimetricToPixel(objPicture.Height, Screen.TwipsPerPixelY)
End If

' Fill surface description
With tSurface
    .lFlags = DDSD_CAPS + DDSD_HEIGHT + DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = iWidth
    .lHeight = iHeight
End With

' Create surface.
Set dixuNewSurface = dixuDirectDraw.CreateSurface(tSurface)

' We put to picture in after for CREATE_FROM_NONE or
' first for the other sources.
If bCreateFrom = CREATE_FROM_NONE Then
    ' Just clear the buffer.
    ClearSurface dixuNewSurface
Else
    ' Put the specified picture onto the surface.
    SetSurface dixuNewSurface, 0, 0, iWidth, iHeight, CREATE_FROM_OLEPICTURE, objPicture
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub SetSurface(dixuSurface As DirectDrawSurface7, ByVal iDestX As Integer, ByVal iDestY As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal bCreateFrom As Byte, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant)

' Purpose: Puts a picture from one of the following
'   sources onto a given surface.  If the given width or
'   height are not the same as the width or height of the
'   picture, then SetSurface will stretch the picture.
'   1)  a resource file
'       bCreateFrom = CREATE_FROM_RES
'       vSource = resource ID
'       vOtherInfo = type of resource (bitmap, icon, cursor)
'   2)  a picture box (viz. HDC)
'       bCreateFrom = CREATE_FROM_HDC
'       vSource = picture box itself
'       vOtherInfo = not used
'   3)  a OLE picture
'       bCreateFrom = CREATE_FROM_OLEPICTURE
'       vSource = StdPicture object
'       vOtherInfo = not used
'   3)  a blank surface
'       bCreateFrom = CREATE_FROM_NONE
'       vSource = not used
'       vOtherInfo = not used
'   NOTE: We NEVER load a picture from a file, the
'       procedure that calls us must do this themselves.
'       This is due to the fact that we cannot take into
'       consideration all possible file formats: it's not
'       worth it.
'   NOTE: NEVER give us a picture that is bigger than
'       the size of the screen.  WE WILL LOCK!

On Error GoTo ErrorHandler

Dim objPicture As StdPicture
Dim hPictureDC As Long
Dim iOriginalWidth As Integer
Dim iOriginalHeight As Integer

' Surface description
Dim tSurface As DDSURFACEDESC2

' Used when stretching
Dim dixuOriginal As DirectDrawSurface7

' Handle on surface context
Dim lSurfaceDC As Long

' Load picture from specified location.
If bCreateFrom = CREATE_FROM_HDC Then
    ' Get the picture DC.
    hPictureDC = vSource.hdc
    ' Copy the bitmap information.
    Set objPicture = vSource.Picture
Else
    Select Case bCreateFrom
        Case CREATE_FROM_RES
            ' Load the picture from the resource file.
            Set objPicture = PictureFromBits(LoadResData(vSource, vOtherInfo))
        Case CREATE_FROM_OLEPICTURE
            ' Just copy the picture from one OLE picture
            ' variable (the one given) to another (ours).
            Set objPicture = vSource
        Case Else
            ' Exit if it's in any other format.
            Exit Sub
    End Select

    ' Create a DC for the picture
    hPictureDC = CreateCompatibleDC(0)

    ' Combine picture and DC
    SelectObject hPictureDC, objPicture.Handle
End If

iOriginalWidth = HimetricToPixel(objPicture.Width, Screen.TwipsPerPixelX)
iOriginalHeight = HimetricToPixel(objPicture.Height, Screen.TwipsPerPixelY)

' If the width and height aren't specified, then set them
' to the actual size of the picture
If iWidth = 0 Then iWidth = iOriginalWidth
If iHeight = 0 Then iHeight = iOriginalHeight

' Fill surface description
With tSurface
    .lFlags = DDSD_CAPS + DDSD_HEIGHT + DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = iOriginalWidth
    .lHeight = iOriginalHeight
End With

' Create surface
Set dixuOriginal = dixuDirectDraw.CreateSurface(tSurface)

' Restore the surface
dixuOriginal.restore

' Get the surface's DC
lSurfaceDC = dixuOriginal.GetDC

BitBlt lSurfaceDC, 0, 0, iOriginalWidth, iOriginalHeight, hPictureDC, 0, 0, vbSrcCopy

dixuOriginal.ReleaseDC lSurfaceDC

' Delete the picture; we've just copied it to a surface.
If bCreateFrom <> CREATE_FROM_HDC Then DeleteDC hPictureDC
Set objPicture = Nothing

Blt dixuSurface, iDestX, iDestY, iWidth, iHeight, dixuOriginal, 0, 0, iOriginalWidth, iOriginalHeight, False

Set dixuOriginal = Nothing

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub Flip()

' Purpose: Flips the front surface and back surface.

On Error GoTo ErrorHandler

dixuFrontSurface.Flip Nothing, DDFLIP_WAIT

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub Blt(ByVal dixuDest As DirectDrawSurface7, ByVal iDestX As Integer, ByVal iDestY As Integer, ByVal iDestWidth As Integer, ByVal iDestHeight As Integer, ByVal dixuSource As DirectDrawSurface7, ByVal iSourceX As Integer, ByVal iSourceY As Integer, ByVal iSourceWidth As Integer, ByVal iSourceHeight As Integer, ByVal bUseColorKey As Byte)

' Purpose: Copies a surface (or part of it) onto another
'   surface (or part of it).  This procedure supports
'   stretching.  Call BltFast for speed without stretching.

On Error GoTo ErrorHandler

Dim lFlags As Long

' Just use the width of the surfaces if the information
' isn't given.
If iSourceWidth = 0 Then iSourceWidth = Width(dixuSource)
If iSourceHeight = 0 Then iSourceHeight = Height(dixuSource)
If iDestWidth = 0 Then iDestWidth = Width(dixuDest)
If iDestHeight = 0 Then iDestHeight = Width(dixuDest)

lFlags = DDBLT_WAIT

If bUseColorKey = True Then lFlags = lFlags Or DDBLT_KEYSRC

dixuDest.Blt GetRect(iDestX, iDestY, iDestWidth, iDestHeight), dixuSource, GetRect(iSourceX, iSourceY, iSourceWidth, iSourceHeight), lFlags

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub BltFast(ByVal dixuDest As DirectDrawSurface7, ByVal iDestX As Integer, ByVal iDestY As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal dixuSource As DirectDrawSurface7, ByVal iSourceX As Integer, ByVal iSourceY As Integer, ByVal bUseColorKey As Byte)

' Purpose: Copies a surface (or part of it) onto another
'   surface (or part of it).  This procedure DOES NOT
'   support stretching.  Call Blt for stretching.

On Error GoTo ErrorHandler

Dim lFlags As Long

' Just use the width of the surfaces if the information
' isn't given.
If iWidth = 0 Then iWidth = Width(dixuSource)
If iHeight = 0 Then iHeight = Height(dixuSource)

lFlags = DDBLTFAST_WAIT

If bUseColorKey = True Then lFlags = lFlags Or DDBLTFAST_SRCCOLORKEY

dixuDest.BltFast iDestX, iDestY, dixuSource, GetRect(iSourceX, iSourceY, iWidth, iHeight), lFlags

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub ClearSurface(dixuSurface As DirectDrawSurface7, Optional ByVal iLeft As Integer, Optional ByVal iTop As Integer, Optional ByVal iWidth As Integer, Optional ByVal iHeight As Integer)

' Purpose: Clears a region of a surface.

On Error GoTo ErrorHandler

dixuSurface.BltColorFill GetRect(iLeft, iTop, iWidth, iHeight), lBackColor

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub RestoreSurfaces()

' Purpose: Restores all the surfaces to working order.
'   This must be called after the window has been
'   re-maximized from a minimized state.

On Error GoTo ErrorHandler

dixuDirectDraw.RestoreAllSurfaces

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub FlipToGDI()

' Purpose: Flips the Windows GDI surface on the screen,
'   so that Windows stuff can be seen.  Must also be
'   called after a DirectDraw session that included
'   flipping.

On Error GoTo ErrorHandler

dixuDirectDraw.FlipToGDISurface

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub SetColorKey(dixuSurface As DirectDrawSurface7, ByVal lColor As Long)

' Purpose: Sets a specific color to be used as the color key.

On Error GoTo ErrorHandler

Dim tColorKey As DDCOLORKEY

With tColorKey
    .high = lColor
    .low = .high
End With

dixuSurface.SetColorKey DDCKEY_SRCBLT, tColorKey

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Property Get Width(dixuSurface As DirectDrawSurface7) As Integer

' Purpose: Returns the width of a specific surface.

On Error GoTo ErrorHandler

Dim tSurfaceDesc As DDSURFACEDESC2

dixuSurface.GetSurfaceDesc tSurfaceDesc

Width = tSurfaceDesc.lWidth

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Property
Property Get Height(dixuSurface As DirectDrawSurface7) As Integer

' Purpose: Returns the height of a specific surface.

On Error GoTo ErrorHandler

Dim tSurfaceDesc As DDSURFACEDESC2

dixuSurface.GetSurfaceDesc tSurfaceDesc

Height = tSurfaceDesc.lHeight

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Property
Public Function TestState() As Byte

' Purpose: Makes sure we still have access to the screen.

On Error GoTo ErrorHandler

If dixuDirectDraw.TestCooperativeLevel = DD_OK Then
    TestState = True
Else: TestState = False
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function GetRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer) As RECT

' Purpose: Returns a rect type with the given info.

On Error GoTo ErrorHandler

With GetRect
    .Left = Left
    .Top = Top
    .Right = .Left + Width
    .Bottom = .Top + Height
End With

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function HimetricToPixel(ByVal iValue As Long, ByVal iTwipsPerPixel As Integer) As Integer

' Purpose: Converts a Himetric dimension into pixel format.

On Error GoTo ErrorHandler

HimetricToPixel = CLng((iValue * 0.001) * 567 / iTwipsPerPixel)

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub DrawText(dixuSurface As DirectDrawSurface7, ByVal sText As String, ByVal bCenter As Byte, Optional ByVal X As Integer, Optional ByVal Y As Integer)

On Error GoTo ErrorHandler

Dim lSurfaceDC As Long
Dim tTextSize As SIZE

If bCenter = True Then
    lSurfaceDC = dixuSurface.GetDC

    GetTextExtentPoint32 lSurfaceDC, sText, Len(sText), tTextSize

    dixuSurface.ReleaseDC lSurfaceDC

    X = (Width(dixuSurface) - tTextSize.cx) \ 2
    Y = (Height(dixuSurface) - tTextSize.cy) \ 2
End If

dixuSurface.DrawText X, Y, sText, False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub MaskToShiftValues(ByVal Mask As Long, ShiftRight As Long, ShiftLeft As Long)

' Purpose: Used to shift a value in order to calculate
'   the DirectDraw version of a RGB color value.

On Error GoTo ErrorHandler

Dim ZeroBitCount As Long
Dim OneBitCount As Long

' Count zero bits
ZeroBitCount = 0
Do While (Mask And 1) = 0
    ZeroBitCount = ZeroBitCount + 1
    Mask = Mask \ 2 ' Shift right
Loop

' Count one bits
OneBitCount = 0
Do While (Mask And 1) = 1
    OneBitCount = OneBitCount + 1
    Mask = Mask \ 2 ' Shift right
Loop

' Shift right 8-OneBitCount bits
ShiftRight = 2 ^ (8 - OneBitCount)
' Shift left ZeroBitCount bits
ShiftLeft = 2 ^ ZeroBitCount

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function TranslateColor(ByVal lColor As Long) As Long

' Purpose: Used internally to translate a specific color
'   into the DirectDraw format.

On Error GoTo ErrorHandler

If RedShiftRight <> 0 And GreenShiftRight <> 0 And BlueShiftRight <> 0 Then
    TranslateColor = (ExtractColor(lColor, 1) \ RedShiftRight) * RedShiftLeft + (ExtractColor(lColor, 2) \ GreenShiftRight) * GreenShiftLeft + (ExtractColor(lColor, 3) \ BlueShiftRight) * BlueShiftLeft
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Function ExtractColor(lRGB As Long, iBytePos As Integer)

' Purpose: Extract a specific color element from an
'   RGB value (1 = Red; 2 = Blue; 3 = Green).

On Error GoTo ErrorHandler

Dim Byte1 As Integer, Byte2 As Integer, Byte3 As Integer, Byte4 As Integer
Dim Sub2 As Long, Sub3 As Long, Sub4 As Long

Byte4 = lRGB \ &H1000000
Sub4 = Byte4 * &H1000000

Byte3 = (lRGB - Sub4) \ &H10000
Sub3 = Byte3 * &H10000

Byte2 = (lRGB - Sub4 - Sub3) \ &H100
Sub2 = Byte2 * CLng(&H100)

Byte1 = lRGB - Sub4 - Sub3 - Sub2

Select Case iBytePos
    Case 1: ExtractColor = Byte1
    Case 2: ExtractColor = Byte2
    Case 3: ExtractColor = Byte3
    Case 4: ExtractColor = Byte4
End Select

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub KillDirectDraw()

' Purpose: Shuts down all DirectDraw stuff.  We put this
'   here instead of in the Class_Terminate procedure
'   because this way the user can call GetDisplayModes
'   without losing the information later, which would
'   happen if they were to set us to nothing.

On Error GoTo ErrorHandler

If dixuDirectDraw Is Nothing Then Exit Sub

With dixuDirectDraw
    .FlipToGDISurface
    .RestoreDisplayMode
    .SetCooperativeLevel lWndHandle, DDSCL_NORMAL
End With

' It appears that the order is of great importance
Set dixuBackBuffer = Nothing
Set dixuFrontSurface = Nothing
Set dixuDirectDraw = Nothing

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
