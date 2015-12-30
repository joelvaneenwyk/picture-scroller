VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirectInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lWndHandle As Long

Dim diDirectInput As DirectInputA
Dim diKeyboard As DirectInputDeviceA
Dim diMouse As DirectInputDeviceA
Public Function InitDirectInput(ByVal hwnd As Long) As Byte

On Error GoTo ErrorHandler

Dim tGuid As GUID
Dim tDataFormat As DIDATAFORMAT

Dim bKeyboard(0 To 255) As Byte

lWndHandle = hwnd

' Create DirectInput.
DirectInputCreateA App.hInstance, &H500, diDirectInput, Nothing

' Define system keyboard guid.
With tGuid
    .Data1 = &H6F1D2B61
    .Data2 = &HD5A0
    .Data3 = &H11CF
    .Data4(0) = &HBF
    .Data4(1) = &HC7
    .Data4(2) = &H44
    .Data4(3) = &H45
    .Data4(4) = &H53
    .Data4(5) = &H54
End With

' Create keyboard device.
diDirectInput.CreateDevice tGuid, diKeyboard, Nothing

' Set cooperative level to be foreground only and
' non exclusive.
diKeyboard.SetCooperativeLevel lWndHandle, DISCL_FOREGROUND Or DISCL_NONEXCLUSIVE

InitDirectInput = True

Exit Function

ErrorHandler:
On Error GoTo 0
InitDirectInput = False

End Function
Private Sub Class_Terminate()

Set diKeyboard = Nothing

End Sub
