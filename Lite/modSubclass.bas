Attribute VB_Name = "modSubclass"
Option Explicit

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
Public Function MenuProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Purpose: This is used to detect when the mouse is
'   over a menu item.

On Error GoTo ErrorHandler

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

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
