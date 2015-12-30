Attribute VB_Name = "modTimeSavers"
Option Explicit

Public bIsWinNT As Byte

Private Const HH_CLOSE_ALL = &H12
Private Const HH_INITIALIZE = &H1C
Private Const HH_UNINITIALIZE = &H1D

Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function HtmlHelpA Lib "hhctrl.ocx" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
' --------------------------------------
' Set a window to "AlwaysOnTop"

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const conHwndTopmost = -1
Private Const conSwpShowWindow = &H40

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function LastSlash(ByVal sPath As String) As Integer

' Purpose: This function finds the position of the
'   last slash in a path.  This number can then be
'   used to extract the file name or the path of
'   a file.

On Error GoTo ErrorHandler

Dim iPosition As Integer
Dim iLastPosition As Integer

Do
    iPosition = InStr(iPosition + 1, sPath, "\")

    If iPosition <> 0 Then
        iLastPosition = iPosition
    Else: Exit Do
    End If
Loop

LastSlash = iLastPosition

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function TrimNulls(ByVal sText As String) As String

' Purpose: Trims the null characters from a string.

On Error GoTo ErrorHandler

Dim iPosition As Integer

iPosition = InStr(sText, vbNullChar)

If iPosition <> 0 Then
    TrimNulls = Left$(sText, iPosition - 1)
Else: TrimNulls = sText
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub WaitProcess(ByVal sWords As String, ByVal bCancelEnabled As Byte, Optional frmOwner As Form)

' Purpose: Automates the process of showing the wait
'   form so that the user is appeased

On Error GoTo ErrorHandler

Screen.MousePointer = 11

With frmWait
    ' Set the text that is shown in the label
    .sWords = sWords

    ' Set whether or not the cancel button can be clicked
    .bCancelEnabled = bCancelEnabled

    If Not frmOwner Is Nothing Then
        ' Show the wait window
        frmOwner.Enabled = False

        .Show , frmOwner
    Else: .Show
    End If

    DoEvents
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub EndWaitProcess(Optional frmOwner As Form)

' Purpose: Shutdown a previously started wait process.

On Error GoTo ErrorHandler

If Not frmOwner Is Nothing Then
    frmOwner.Enabled = True
    frmOwner.Show
    DoEvents
End If

Unload frmWait
DoEvents

Screen.MousePointer = 0

' If this function has been called, then it means that
' everyone knows that they canceled or the process
' was completed, so we can now reset bCancelOp.
bCancelOp = False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Function NormalizePath(sPath As String) As String

' Purpose: Add a slash at the end of the path if
'   necessary.  Returns the completed path, while also
'   changing sPath.

On Error GoTo ErrorHandler

If GetAttr(sPath) And vbDirectory Then
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"

    NormalizePath = sPath
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function FileExists(ByVal sPath As String) As Boolean

' Purpose: Checks if a file exists.

On Error GoTo ErrorHandler

If bIsWinNT = True Then
    FileExists = SHFileExists(StrConv(sPath, vbUnicode))
Else
    FileExists = SHFileExists(sPath)
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function IsWinNT() As Boolean

' Purpose: Detect whether this is Window NT or not.

On Error GoTo ErrorHandler

Dim tOSVersion As OSVERSIONINFO

tOSVersion.dwOSVersionInfoSize = Len(tOSVersion)
GetVersionEx tOSVersion

IsWinNT = (tOSVersion.dwPlatformId = VER_PLATFORM_WIN32_NT)

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub SetDefaults()

On Error GoTo ErrorHandler

Dim nIndex As Integer

With tProgramOptions
    .iInterval = 3000
    .bPictureSize = 1

    .lInfoColor = RGB(255, 255, 255)
    ReDim .bCustomColors(0 To 16 * 4 - 1) As Byte

    For nIndex = 0 To UBound(.bCustomColors)
        .bCustomColors(nIndex) = 255
    Next nIndex

    .bSoundEffects = 1

    ReDim .bTransitions(1 To NUM_OF_TRANSITIONS)

    For nIndex = 1 To NUM_OF_TRANSITIONS
        .bTransitions(nIndex) = nIndex
    Next nIndex

    .MAX_LEN = 1000
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub SaveUserOptions()

On Error GoTo ErrorHandler

' Now that we're saving options, next time won't be a
' first start.
tProgramOptions.bNotFirstStart = True

' Save the user's options
Open sAppPath & "Options.dat" For Binary As FILENUM_OPTIONS
Put FILENUM_OPTIONS, , tProgramOptions
Close FILENUM_OPTIONS

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub AlwaysOnTop(ByVal hwnd As Long)

SetWindowPos hwnd, conHwndTopmost, 0, 0, 0, 0, conSwpShowWindow Or SWP_NOMOVE Or SWP_NOSIZE

End Sub
Public Sub ShowHelp(ByVal hwnd As Long, ByVal sTopic As String, ByVal bSound As Byte)

On Error Resume Next

Static bStarted As Byte

If bStarted = False Then
    HtmlHelpA 0, 0, HH_INITIALIZE, ""
    bStarted = True
End If

If sTopic = "<<CLOSEALL>>" Then
    If bStarted = True Then
        HtmlHelpA 0, "", HH_CLOSE_ALL, ""
        HtmlHelpA 0, 0, HH_UNINITIALIZE, ""
    End If

    Exit Sub
End If

' Play help message if sound effects are enabled.
If tProgramOptions.bSoundEffects = 1 And DirectSound.bInitOK = True And bSound = True Then
    DirectSound.PlaySound "HELP", False
End If

If HtmlHelpA(hwnd, sAppPath & "Picture Scroller.chm", 0, ByVal sTopic) = False Then
    MsgBox "Unable to load Picture Scroller's Help file.  To view the help file, you may need to install HTML Help 1.1 or greater, or you may need to reinstall Picture Scroller.  View the 'readme.txt' file for more details.", vbCritical
End If

End Sub
Public Function GetProportional(ByVal iWidth As Variant, ByVal iHeight As Variant, iAimWidth As Variant, iAimHeight As Variant) As Single

On Error GoTo ErrorHandler

If iWidth / iAimWidth > iHeight / iAimHeight Then
    ' If the width is greater than the height, then change
    ' the height.
    iAimHeight = (iAimWidth * iHeight) / iWidth
ElseIf iHeight / iAimHeight > iWidth / iAimWidth Then
    ' If the height is greater than the width, then change
    ' the width.
    iAimWidth = (iAimHeight * iWidth) / iHeight
End If

GetProportional = iAimHeight / iHeight

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Sub SetMouseCount(ByVal bOriginal As Byte)

On Error GoTo ErrorHandler

Static iMouseCount As Integer

If bOriginal = False Then
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
Else
    ' Set the mouse count to the original count
    Do While ShowCursor(False) >= iMouseCount
    Loop
    Do While ShowCursor(True) <= iMouseCount
    Loop

    iMouseCount = 0
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub ErrHandle()

End Sub
