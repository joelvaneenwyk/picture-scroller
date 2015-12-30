Attribute VB_Name = "modSubclass"
Option Explicit

Const WM_ACTIVATEAPP = &H1C
Const WM_SIZE = &H5

Const MM_MCINOTIFY = &H3B9
Const MCI_NOTIFY_SUCCESSFUL = &H1

Const WM_MENUSELECT = &H11F

Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

' Holds the Visual Basic procedure identifier.
' We call this so that VB can handle a message after
' we have done what we want with it.
Public lPrevWndProc As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Function ScrollingProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Purpose: This is used when frmScrolling is in process.

Dim LoWord As Integer
Dim HiWord As Integer

If Msg = WM_ACTIVATEAPP Then
    ' Set bMinimized to True if the user just minimized us
    If wParam = False Then frmScroller.Minimize True

    ' Call the regular Visual Basic procedure, but
    ' return zero
    ScrollingProc = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
    ScrollingProc = 0
ElseIf Msg = WM_SIZE Then
    GetHiLoWord lParam, LoWord, HiWord

    ' If the user has just reactivated us to fullscreen,
    ' then set bMinimized to False, so drawing can continue
    If LoWord = iScreenWidth And HiWord = iScreenHeight Then
        frmScroller.Minimize False
    End If

    ' Call the regular Visual Basic procedure, but
    ' return zero
    ScrollingProc = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
    ScrollingProc = 0
ElseIf Msg = MM_MCINOTIFY Then
    ' Notify the CallBack sub of BackMusic, so that
    ' they can do whatever they need to do.
    If wParam = MCI_NOTIFY_SUCCESSFUL Then BackMusic.CallBack

    ScrollingProc = 0
Else
    ' If this message isn't something we want to deal
    ' with, then give it to the Visual Basic procedure
    ScrollingProc = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
End If

End Function
Public Function MenuProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Purpose: This is used to detect when the mouse is
'   over a menu item.

Dim lReturnVal As Long

If Msg = WM_MENUSELECT Then
    frmMain.MenuSelect hwnd

    ' Call the regular Visual Basic procedure, but
    ' return zero
    MenuProc = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
    MenuProc = 0
Else
    ' If this message isn't something we want to deal
    ' with, then give it to the Visual Basic procedure
    MenuProc = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
End If

End Function
'Public Function DisplayModesProc(lpDDSurfaceDesc As DDSURFACEDESC2, ByVal lpContext As Long) As Long
'
'' Purpose: This is called by DirectDraw when we are
''   enumerating the display modes.  We, in turn, call
''   the DirectDraw class so they can do whatever.
'
'DirectDraw.ModesCallback lpDDSurfaceDesc
'
'' Tell DirectDraw to keep it coming.
'DisplayModesProc = DDENUMRET_OK
'
'End Function
Private Sub GetHiLoWord(lWord As Long, LoWord As Integer, HiWord As Integer)

' Purpose: Extract the high and low word value
'   from the long value given

LoWord = CInt(lWord And &HFFFF&)
HiWord = CInt(lWord \ &H10000)

End Sub
