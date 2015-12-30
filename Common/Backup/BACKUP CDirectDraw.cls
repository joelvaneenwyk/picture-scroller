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

Private Enum CBoolean
    CFalse = 0
    CTrue = 1
End Enum

Private Type GUID
    dwData1 As Long
    wData2 As Integer
    wData3 As Integer
    abData4(7) As Byte
End Type

Private Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Private Const GMEM_MOVEABLE = &H2

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As CBoolean, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As CBoolean, riid As GUID, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const conHwndTopmost = -1
Private Const conSwpShowWindow = &H40

Private Const DT_LEFT = &H0

Private Const LOGPIXELSY = 90
Private Const LOGPIXELSX = 88

Private Const FF_SCRIPT = 64

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

' These functions may need to be public later
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal nWeight As Long, ByVal bItalic As Long, ByVal bUnderline As Long, ByVal bStrikeout As Long, ByVal lCharSet As Long, ByVal lOutPrecision As Long, ByVal lClipPrecision As Long, ByVal lQuality As Long, ByVal lPitchFamily As Long, ByVal lFace As String) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

' Handle of the window owning us.
Dim lWndHandle As Long

' Default and necessary DirectDraw surfaces
Dim dixuDirectDraw As DirectDraw7
Dim dixuFrontSurface As DirectDrawSurface7
Dim dixuBackBuffer As DirectDrawSurface7

' Collection of user created surfaces
Dim colSurfaces As New Collection

' Holds a list of supported display modes.
Dim tDisplayModes() As DISPLAY_MODE

' TRUE if we're using DirectDraw's "Flip" function
Dim bFlip As Byte

Dim RedShiftLeft As Long
Dim RedShiftRight As Long
Dim GreenShiftLeft As Long
Dim GreenShiftRight As Long
Dim BlueShiftLeft As Long
Dim BlueShiftRight As Long

Dim lBackColor As Long
Dim lTextColor As Long
Public Function InitDirectDraw(ByVal hwnd As Long, ByVal bTrueFlip As Byte, ByVal bPalettized As Byte) As Byte

' Purpose: Initializes DirectDraw.  This includes creating
'   a front surface and getting the back surface.
'   bTrueFlip
'       TRUE - we use the true DirectDraw function "Flip"
'       FALSE - we emulate that function by using "BltFast"
'   bPalettized
'       TRUE - sets the display mode at the current
'              resolution but with 8 bpp.
'       FALSE - uses whatever is set in tProgramOptions

On Error GoTo ErrorHandler

Dim tNewDisplayMode As DISPLAY_MODE
Dim tSurface As DDSURFACEDESC2
Dim tBlankSurface As DDSURFACEDESC2
Dim tSurfaceCaps As DDSCAPS2

lWndHandle = hwnd

SetWindowPos lWndHandle, conHwndTopmost, 0, 0, 0, 0, conSwpShowWindow Or SWP_NOMOVE Or SWP_NOSIZE

Set dixuDirectDraw = Nothing

Set dixuDirectDraw = DirectX.DirectDrawCreate("")
dixuDirectDraw.SetCooperativeLevel lWndHandle, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT

' If we're NOT to run in palettized mode, then
' use whatever the user set.
If bPalettized = False Then
    Select Case tProgramOptions.bScreenSetting
        Case 0
            ' If the user just wanted to keep the current
            ' resolution, then just set the variable so
            ' that we know.
            With tNewDisplayMode
                .iWidth = Screen.Width \ Screen.TwipsPerPixelX
                .iHeight = Screen.Height \ Screen.TwipsPerPixelY
            End With
        Case 1
            ' Just select the first display mode; it will be
            ' the fastest.
            tNewDisplayMode = tDisplayModes(0)
        Case 2
            ' Just select the last display mode; it will have
            ' the best quality.
            tNewDisplayMode = tDisplayModes(UBound(tDisplayModes))
        Case 3
            tNewDisplayMode = tProgramOptions.tDisplayMode
    End Select
Else
    ' Otherwise, use the current resolution with 8 bpp.
    With tNewDisplayMode
        .iWidth = 640
        .iHeight = 480
        .bBPP = 8
    End With
End If

With tNewDisplayMode
    ' Set the screen resolution, if the user has choosen
    ' to do so or if bPalettized is TRUE.
    If tProgramOptions.bScreenSetting <> 0 Or bPalettized = True Then
        dixuDirectDraw.SetDisplayMode .iWidth, .iHeight, .bBPP, 0, 0
    End If

    iScreenWidth = .iWidth
    iScreenHeight = .iHeight
End With

' Set the properties of the front surface, depending
' on whether where creating a true flippable surface.
If bTrueFlip = True Then
    With tSurface
        .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
        .lBackBufferCount = 1
    End With
Else
    With tSurface
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
End If

' Create the front surface
Set dixuFrontSurface = dixuDirectDraw.CreateSurface(tSurface)

dixuFrontSurface.restore

' Create a different type of back buffer depending on
' bTrueFlip.
If bTrueFlip = True Then
    ' Create a true flippable surface.
    tSurfaceCaps.lCaps = DDSCAPS_BACKBUFFER
    Set dixuBackBuffer = dixuFrontSurface.GetAttachedSurface(tSurfaceCaps)
Else
    ' Create a regular surface, where we emulate flipping.
    tSurface = tBlankSurface

    ' Fill surface description
    With tSurface
        .lFlags = DDSD_CAPS + DDSD_HEIGHT + DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = iScreenWidth
        .lHeight = iScreenHeight
    End With

    ' Create surface.
    Set dixuBackBuffer = dixuDirectDraw.CreateSurface(tSurface)
End If

dixuBackBuffer.restore

Me.ClearRegion "BACKBUFFER"

dixuFrontSurface.GetSurfaceDesc tSurface

' Set the background color.
If tSurface.ddpfPixelFormat.lRGBBitCount > 8 Then
    MaskToShiftValues tSurface.ddpfPixelFormat.lRBitMask, RedShiftRight, RedShiftLeft
    MaskToShiftValues tSurface.ddpfPixelFormat.lGBitMask, GreenShiftRight, GreenShiftLeft
    MaskToShiftValues tSurface.ddpfPixelFormat.lBBitMask, BlueShiftRight, BlueShiftLeft

    lBackColor = TranslateColor(tProgramOptions.lBackColor)
Else: lBackColor = tProgramOptions.lBackColor
End If

' Store the text color.
lTextColor = tProgramOptions.lInfoColor

' Set the font color and use transparent fonts.
dixuBackBuffer.SetForeColor lTextColor
dixuBackBuffer.SetFontTransparency True

' Store exactly what mode we are in.
bFlip = bTrueFlip

InitDirectDraw = True

Exit Function

ErrorHandler:
On Error GoTo 0
InitDirectDraw = False

End Function
Public Sub GetDisplayModes(ByVal hwnd As Long)

' Purpose: Retrieves a list of all the supported display
'   modes on this computer.  It does this by telling
'   DirectDraw to call us back for each display mode, by
'   calling ModesCallback.

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

End Sub
Property Get ModeWidth(ByVal iIndex As Integer) As Integer

' Purpose: Returns the width of a specific display mode.

ModeWidth = tDisplayModes(iIndex).iWidth

End Property
Property Get ModeCount() As Integer

' Purpose: Returns the number of supported display modes.

ModeCount = UBound(tDisplayModes)

End Property
Property Get ModeHeight(ByVal iIndex As Integer) As Integer

' Purpose: Returns the height of a specific display mode.

ModeHeight = tDisplayModes(iIndex).iHeight

End Property
Property Get ModeBPP(ByVal iIndex As Integer) As Byte

' Purpose: Returns the number of colors of the specified
'   display mode.

ModeBPP = tDisplayModes(iIndex).bBPP

End Property
Public Sub CreateSurface(ByVal sKey As String, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal bCreateFrom As CREATE_METHODS, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant)

' Purpose: Creates a surface of the specified dimensions.
'   If bCreateFrom is not CREATE_FROM_NONE, then we also
'   load the picture from the source into objPicture.
'   If no dimensions are specified, we can take the sizes
'   from objPicture.  Again, if bCreateFrom is not
'   CREATE_FROM_NONE, we pass what we know (objPicture)
'   to SetPicture so that they can put the picture on
'   the surface.

Dim objPicture As New StdPicture

' Surface description
Dim tSurface As DDSURFACEDESC2
Dim dixuNewSurface As DirectDrawSurface7
Dim objSurface As New CSurface

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

' Set the font color and use transparent fonts.
dixuNewSurface.SetForeColor lTextColor
dixuNewSurface.SetFontTransparency True

' We put to picture in after for CREATE_FROM_NONE or
' first for the other sources.
If bCreateFrom = CREATE_FROM_NONE Then
    ' Just clear the buffer.
    ClearBuffer dixuNewSurface

    GoSub SetSurfaceObject
Else
    GoSub SetSurfaceObject

    ' Put the specified picture onto the surface.
    SetSurface sKey, 0, 0, iWidth, iHeight, CREATE_FROM_OLEPICTURE, objPicture
End If

Set dixuNewSurface = Nothing
Set objSurface = Nothing

Exit Sub

SetSurfaceObject:
' Put the information in the surface class.
With objSurface
    Set .Surface = dixuNewSurface
    .Width = iWidth
    .Height = iHeight
End With

' Add the surface to the collection.
colSurfaces.Add objSurface, sKey

Return

End Sub
Public Sub SetSurface(ByVal sKey As String, ByVal iDestX As Integer, ByVal iDestY As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal bCreateFrom As Byte, ByVal vSource As Variant, Optional ByVal vOtherInfo As Variant)

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

' Used for stretching
Dim tDestRect As RECT
Dim tSourceRect As RECT

Dim dixuSurface As New CSurface

Set dixuSurface = colSurfaces(sKey)

dixuSurface.Surface.restore

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

With tDestRect
    .Left = iDestX
    .Top = iDestY
    .Right = .Left + iWidth
    .Bottom = .Top + iHeight
End With

With tSourceRect
    .Right = iOriginalWidth
    .Bottom = iOriginalHeight
End With

dixuSurface.Surface.Blt tDestRect, dixuOriginal, tSourceRect, DDBLT_WAIT

Set dixuOriginal = Nothing
Set dixuSurface = Nothing

End Sub
Public Sub Flip()

' Purpose: Flips the front surface and back surface.

Dim tWholeSurface As RECT

If bFlip = True Then
    dixuFrontSurface.Flip Nothing, DDFLIP_WAIT
Else: dixuFrontSurface.BltFast 0, 0, dixuBackBuffer, tWholeSurface, DDBLTFAST_WAIT
End If

End Sub
Public Sub Blt(ByVal sDestKey As String, ByVal iDestX As Integer, ByVal iDestY As Integer, ByVal iDestWidth As Integer, ByVal iDestHeight As Integer, ByVal sSourceKey As String, ByVal iSourceX As Integer, ByVal iSourceY As Integer, ByVal iSourceWidth As Integer, ByVal iSourceHeight As Integer, ByVal bUseColorKey As Byte)

' Purpose: Copies a surface (or part of it) onto another
'   surface (or part of it).  This procedure supports
'   stretching.  Call BltFast for speed without stretching.

Dim tSourceRect As RECT
Dim tDestRect As RECT
Dim dixuSourceSurface As New CSurface
Dim dixuDestSurface As New CSurface
Dim lFlags As Long
Dim tBltFx As DDBLTFX

Set dixuSourceSurface = colSurfaces(sSourceKey)

dixuSourceSurface.Surface.restore

If sDestKey <> "BACKBUFFER" Then
    Set dixuDestSurface = colSurfaces(sDestKey)

    dixuDestSurface.Surface.restore
End If

If iSourceWidth = 0 Then iSourceWidth = Width(sSourceKey)
If iSourceHeight = 0 Then iSourceHeight = Height(sSourceKey)

If iDestWidth = 0 Then iDestWidth = iSourceWidth
If iDestHeight = 0 Then iDestHeight = iSourceHeight

With tSourceRect
    .Left = iSourceX
    .Top = iSourceY
    .Right = .Left + iSourceWidth
    .Bottom = .Top + iSourceHeight
End With

With tDestRect
    .Left = iDestX
    .Top = iDestY
    .Right = .Left + iDestWidth
    .Bottom = .Top + iDestHeight
End With

lFlags = DDBLT_WAIT

If bUseColorKey = True Then
    lFlags = lFlags Or DDBLT_KEYSRC
End If

If sDestKey = "BACKBUFFER" Then
    dixuBackBuffer.Blt tDestRect, dixuSourceSurface.Surface, tSourceRect, lFlags
Else
    dixuDestSurface.Surface.Blt tDestRect, dixuSourceSurface.Surface, tSourceRect, lFlags

    Set dixuDestSurface = Nothing
End If

Set dixuSourceSurface = Nothing

End Sub
Public Sub BltFast(ByVal sDestKey As String, ByVal iDestX As Integer, ByVal iDestY As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal sSourceKey As String, ByVal iSourceX As Integer, ByVal iSourceY As Integer, ByVal bUseColorKey As Byte)

' Purpose: Copies a surface (or part of it) onto another
'   surface (or part of it).  This procedure DOES NOT
'   support stretching.  Call Blt for stretching.

Dim tSourceRect As RECT
Dim dixuSourceSurface As New CSurface
Dim dixuDestSurface As New CSurface
Dim lFlags As Long

Set dixuSourceSurface = colSurfaces(sSourceKey)

dixuSourceSurface.Surface.restore

If sDestKey <> "BACKBUFFER" Then
    Set dixuDestSurface = colSurfaces(sDestKey)

    dixuDestSurface.Surface.restore
End If

If iWidth = 0 Then iWidth = Width(sSourceKey)
If iHeight = 0 Then iHeight = Height(sSourceKey)

With tSourceRect
    .Left = iSourceX
    .Top = iSourceY
    .Right = .Left + iWidth
    .Bottom = .Top + iHeight
End With

lFlags = DDBLTFAST_WAIT

If bUseColorKey = True Then
    lFlags = lFlags Or DDBLTFAST_SRCCOLORKEY
End If

If sDestKey = "BACKBUFFER" Then
    dixuBackBuffer.BltFast iDestX, iDestY, dixuSourceSurface.Surface, tSourceRect, lFlags
Else
    dixuDestSurface.Surface.BltFast iDestX, iDestY, dixuSourceSurface.Surface, tSourceRect, lFlags

    Set dixuDestSurface = Nothing
End If

Set dixuSourceSurface = Nothing

End Sub
Public Sub ClearRegion(ByVal sKey As String, Optional ByVal iLeft As Integer, Optional ByVal iTop As Integer, Optional ByVal iWidth As Integer, Optional ByVal iHeight As Integer)

' Purpose: Clears a surface (or portion of it).

Dim dixuSurface As New CSurface
Dim tRect As RECT
Dim bRegion As Byte

If iWidth <> 0 And iHeight <> 0 Then
    With tRect
        .Left = iLeft
        .Right = .Left + iWidth
        .Top = iTop
        .Bottom = .Top + iHeight
    End With

    bRegion = True
End If

Select Case sKey
    Case "BACKBUFFER"
        If bRegion = True Then
            dixuBackBuffer.BltColorFill tRect, lBackColor
        Else: ClearBuffer dixuBackBuffer
        End If
    Case Else
        Set dixuSurface = colSurfaces(sKey)

        dixuSurface.Surface.restore

        If bRegion = True Then
            dixuSurface.Surface.BltColorFill tRect, lBackColor
        Else: ClearBuffer dixuSurface.Surface
        End If

        Set dixuSurface = Nothing
End Select

End Sub
Public Sub RemoveSurface(ByVal sKey As String)

' Purpose: Attempts to remove a surface from the
'   collection of surfaces.  This will not give an error
'   even if the surface does not exist.

On Error Resume Next

colSurfaces.Remove sKey

End Sub
Public Sub RestoreSurfaces()

' Purpose: Restores all the surfaces to working order.
'   This must be called after the window has been
'   re-maximized from a minimized state.

dixuDirectDraw.RestoreAllSurfaces

End Sub
Public Sub FlipToGDI()

' Purpose: Flips the Windows GDI surface on the screen,
'   so that Windows stuff can be seen.  Must also be
'   called after a DirectDraw session that included
'   flipping.

dixuDirectDraw.FlipToGDISurface

End Sub
Public Sub SetColorKey(ByVal sKey As String, ByVal lColor As Long)

' Purpose: Sets a specific color to be used as the color key.

Dim tColorKey As DDCOLORKEY
Dim dixuSurface As New CSurface

With tColorKey
    .high = lColor
    .low = .high
End With

Set dixuSurface = colSurfaces(sKey)

dixuSurface.Surface.restore

dixuSurface.Surface.SetColorKey DDCKEY_SRCBLT, tColorKey

Set dixuSurface = Nothing

End Sub
Public Sub SetFontColor(ByVal sKey As String, ByVal lColor As Long)

' Purpose: Sets the text color.

Dim dixuSurface As New CSurface

If sKey = "BACKBUFFER" Then
    dixuBackBuffer.SetForeColor lColor
Else
    Set dixuSurface = colSurfaces(sKey)
    dixuSurface.Surface.restore

    dixuSurface.Surface.SetForeColor lColor

    Set dixuSurface = Nothing
End If

End Sub
Public Sub SetFont(ByVal sKey As String, ByVal objFont As IFont)

' Purpose: Sets the font of a surface.

Dim dixuSurface As New CSurface

If sKey = "BACKBUFFER" Then
    dixuBackBuffer.SetFont objFont
Else
    Set dixuSurface = colSurfaces(sKey)
    dixuSurface.Surface.restore

    dixuSurface.Surface.SetFont objFont

    Set dixuSurface = Nothing
End If

End Sub
Property Get Width(ByVal sKey As String) As Integer

' Purpose: Returns the width of a specific surface.

Dim dixuSurface As New CSurface

Set dixuSurface = colSurfaces(sKey)

Width = dixuSurface.Width

Set dixuSurface = Nothing

End Property
Property Get Height(ByVal sKey As String) As Integer

' Purpose: Returns the height of a specific surface.

Dim dixuSurface As New CSurface

Set dixuSurface = colSurfaces(sKey)

Height = dixuSurface.Height

Set dixuSurface = Nothing

End Property
Public Sub DrawText(ByVal sKey As String, ByVal sText As String, ByVal X As Integer, ByVal Y As Integer)

' Purpose: Draws text onto the screen.

Dim dixuSurface As New CSurface

If sKey = "BACKBUFFER" Then
    dixuBackBuffer.DrawText X, Y, sText, False
Else
    Set dixuSurface = colSurfaces(sKey)
    dixuSurface.Surface.restore

    dixuSurface.Surface.DrawText X, Y, sText, False

    Set dixuSurface = Nothing
End If

End Sub
Public Function TestState() As Byte

' Purpose: Makes sure we still have access to the screen.

If dixuDirectDraw.TestCooperativeLevel = DD_OK Then
    TestState = True
Else: TestState = False
End If

End Function
Public Sub SetPalette(tPalette() As PALETTEENTRY)

' Purpose: Set the palette of a surface.

Dim dixuPalette As DirectDrawPalette

Set dixuPalette = dixuDirectDraw.CreatePalette(DDPCAPS_8BIT Or DDPCAPS_ALLOW256, tPalette)

dixuFrontSurface.SetPalette dixuPalette

Set dixuPalette = Nothing

End Sub
Public Sub LockSurface(ByVal sKey As String, ByVal X As Integer, ByVal Y As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer)

' Purpose: Locks a surface.  Must be accompanied by
'   an unlock.

Dim dixuSurface As New CSurface
Dim tSurfaceDesc As DDSURFACEDESC2
Dim tRect As RECT

With tRect
    .Left = X
    .Top = Y
    .Right = iWidth - .Left
    .Bottom = iHeight - .Top
End With

If sKey = "BACKBUFFER" Then
    dixuBackBuffer.Lock tRect, tSurfaceDesc, DDLOCK_WAIT, 0
Else
    Set dixuSurface = colSurfaces(sKey)
    dixuSurface.Surface.restore

    dixuSurface.Surface.Lock tRect, tSurfaceDesc, DDLOCK_WAIT, 0
End If

End Sub
Public Sub UnlockSurface(ByVal sKey As String, ByVal X As Integer, ByVal Y As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer)

' Purpose: Unlocks a surface after a lock call.

Dim dixuSurface As New CSurface
Dim tSurfaceDesc As DDSURFACEDESC2
Dim tRect As RECT

With tRect
    .Left = X
    .Top = Y
    .Right = iWidth - .Left
    .Bottom = iHeight - .Top
End With

If sKey = "BACKBUFFER" Then
    dixuBackBuffer.Unlock tRect
Else
    Set dixuSurface = colSurfaces(sKey)
    dixuSurface.Surface.restore

    dixuSurface.Surface.Unlock tRect
End If

End Sub
Public Function GetLockedArray(ByVal sKey As String, bPixels() As Byte)

' Purpose: Returns an array of all the pixels on a
'   locked surface.

Dim dixuSurface As New CSurface

If sKey = "BACKBUFFER" Then
    dixuBackBuffer.GetLockedArray bPixels()
Else
    Set dixuSurface = colSurfaces(sKey)
    dixuSurface.Surface.restore

    dixuSurface.Surface.GetLockedArray bPixels
End If

End Function
Private Function PictureFromBits(abPic() As Byte) As IPicture  ' not a StdPicture!!

' Purpose: Loads a picture from a string of bytes.  Used
'   internally to load a picture from the resource file.

Dim nLow As Long
Dim cbMem  As Long
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown
Dim ipic As IPicture

' Get the size of the picture's bits
On Error GoTo Out

nLow = LBound(abPic)
On Error GoTo 0

cbMem = (UBound(abPic) - nLow) + 1

' Allocate a global memory object
hMem = GlobalAlloc(GMEM_MOVEABLE, cbMem)

If hMem Then
    ' Lock the memory object and get a pointer to it.
    lpMem = GlobalLock(hMem)
    If lpMem Then
        ' Copy the picture bits to the memory pointer and unlock the handle.
        MoveMemory ByVal lpMem, abPic(nLow), cbMem
        Call GlobalUnlock(hMem)

        ' Create an ISteam from the pictures bits (we can explicitly free hMem
        ' below, but we'll have the call do it...)
        If (CreateStreamOnHGlobal(hMem, CTrue, istm) = 0) Then
            If (CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture) = 0) Then
                ' Create an IPicture from the IStream (the docs say the call does not
                ' AddRef its last param, but it looks like the reference counts are correct..)
                Call OleLoadPicture(ByVal ObjPtr(istm), cbMem, CFalse, IID_IPicture, PictureFromBits)
            End If
        End If
    End If

    Call GlobalFree(hMem)
End If

Out:

End Function
Private Sub ClearBuffer(dixuSurface As DirectDrawSurface7)

' Purpose:  Used internally to clear an entire surface.

Dim tWholeSurface As RECT

dixuSurface.BltColorFill tWholeSurface, lBackColor

End Sub
Private Sub MaskToShiftValues(ByVal Mask As Long, ShiftRight As Long, ShiftLeft As Long)

' Purpose: Used to shift a value in order to calculate
'   the DirectDraw version of a RGB color value.

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

End Sub
Private Function TranslateColor(ByVal lColor As Long) As Long

' Purpose: Used internally to translate a specific color
'   into the DirectDraw format.

If RedShiftRight <> 0 And GreenShiftRight <> 0 And BlueShiftRight <> 0 Then
    TranslateColor = (ExtractColor(lColor, 1) \ RedShiftRight) * RedShiftLeft + (ExtractColor(lColor, 2) \ GreenShiftRight) * GreenShiftLeft + (ExtractColor(lColor, 3) \ BlueShiftRight) * BlueShiftLeft
End If

End Function
Private Function ExtractColor(lRGB As Long, iBytePos As Integer)

' Purpose: Extract a specific color element from an
'   RGB value (1 = Red; 2 = Blue; 3 = Green).

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

End Function
Private Function HimetricToPixel(ByVal iValue As Integer, ByVal iTwipsPerPixel As Integer) As Integer

' Purpose: Converts a Himetric dimension into pixel format.

HimetricToPixel = CLng((iValue * 0.001) * 567 / iTwipsPerPixel)

End Function
Public Sub KillDirectDraw()

' Purpose: Shuts down all DirectDraw stuff.  We put this
'   here instead of in the Class_Terminate procedure
'   because this way the user can call GetDisplayModes
'   without losing the information later, which would
'   happen if they were to set us to nothing.

If dixuDirectDraw Is Nothing Then Exit Sub

While colSurfaces.Count > 0
    colSurfaces.Remove 1
Wend

With dixuDirectDraw
    If bFlip = True Then .FlipToGDISurface

    .RestoreDisplayMode
    .SetCooperativeLevel lWndHandle, DDSCL_NORMAL
End With

' It appears that the order is of great importance
Set dixuBackBuffer = Nothing
Set dixuFrontSurface = Nothing
Set dixuDirectDraw = Nothing

End Sub
