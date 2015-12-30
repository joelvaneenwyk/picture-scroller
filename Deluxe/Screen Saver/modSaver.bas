Attribute VB_Name = "modSaver"
Option Explicit

' ----------------------------------------
' Used to run Picture Scroller and wait for it to exit.

Const INFINITE = &HFFFF
Const NORMAL_PRIORITY_CLASS = &H20

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' ----------------------------------------
' Examine the registry for the path to Picture Scroller.

Const HKEY_LOCAL_MACHINE = &H80000002
Const KEY_QUERY_VALUE = &H1

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' ----------------------------------------
' Used to change the password in Windows 95/98.

Private Type CHANGEPWDINFO
    lpUserName As String
    lpPassword As String
    cbPassword As Long
End Type

Private Declare Function PwdChangePassword Lib "mpr.dll" Alias "PwdChangePasswordA" (ByVal lpProvider As String, ByVal hwndOwner As Long, ByVal dwFlags As Long, lpPwdInfo As CHANGEPWDINFO) As Long
' ----------------------------------------
' See if this is Windows NT

Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Sub Main()

' Purpose: This gets run by windows as the screen saver.
'   Examine the command line and do accordingly.

On Error Resume Next

Dim hKey As Long
Dim sProgramPath As String
Dim lPathLen As Long

Dim sCommandLine As String
Dim sRunCommand As String
Dim nIndex As Integer
Dim lHwnd As Long

Dim tPasswordInfo As CHANGEPWDINFO
Dim tProcessInfo As PROCESS_INFORMATION
Dim tStartupInfo As STARTUPINFO
Dim tOSVersion As OSVERSIONINFO

If App.PrevInstance = True Then Exit Sub

' Upper case and trim the command line.
sCommandLine = UCase(Trim(Command))

' Extract the first two characters: they are the
' command we are to do.
sRunCommand = Left(UCase(Trim(Command)), 2)

If sCommandLine = "" Or IsCommand(sRunCommand, "S") = True Then
    ' Run PS in plain Screen Saver run mode.
    sRunCommand = "/SAVER_RUN"
ElseIf IsCommand(sRunCommand, "C") = True Then
    ' Run PS in configuration mode.
    sRunCommand = "/SAVER_CONFIG"
ElseIf IsCommand(sRunCommand, "A") = True Then
    ' We are to change the Screen Saver password.

    ' Get the Windows version information.
    tOSVersion.dwOSVersionInfoSize = Len(tOSVersion)
    GetVersionEx tOSVersion

    ' Only show the "Change Password" window if this isn't
    ' Windows NT/2000.
    If (tOSVersion.dwPlatformId = VER_PLATFORM_WIN32_NT) <> True Then
        ' Extract the hWnd of the window we're to use as the parent.
        For nIndex = 1 To Len(sCommandLine)
            If IsNumeric(Mid(sCommandLine, nIndex, 1)) = True Then
                lHwnd = lHwnd & Mid(sCommandLine, nIndex, 1)
            End If
        Next nIndex

        ' Show Windows' "Change Password" window.
        PwdChangePassword "SCRSAVE", lHwnd, 0, tPasswordInfo
    End If

    ' We're done, so exit.
    Exit Sub
ElseIf IsCommand(sRunCommand, "P") = True Then
    ' Just exit if it's about the preview window.
    Exit Sub
Else
    ' If it's anything else we don't recognize, just
    ' run it as a screen saver.
    sRunCommand = "/SAVER_RUN"
End If

' OK, so now we need the path to PS.
If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Picture Scroller.exe", 0, KEY_QUERY_VALUE, hKey) = 0 Then
    ' Retrieve the length of the program path.
    RegQueryValueEx hKey, "", 0, 0, 0, lPathLen
    If lPathLen = 0 Then GoTo CANNOT_FIND_PROGRAM

    ' Get the program path from the registry.
    sProgramPath = String(lPathLen, 0)
    RegQueryValueEx hKey, "", 0, 0, ByVal sProgramPath, lPathLen

    ' Get rid of any null characters at the end.
    sProgramPath = Left(sProgramPath, lPathLen - 1)
End If

If sProgramPath = "" Then GoTo CANNOT_FIND_PROGRAM
    
' Put quotes around the program path part.
sCommandLine = Chr(34) & sProgramPath & Chr(34) & " " & sRunCommand

' We MUST use this way, because both VB's shell
' function and ShellExecute API both DO NOT
' work.  I think this: VB's shell function
' appears to run the program, but since the
' code continues to run, it seems to immediately
' kill both this program and PS.  ShellExecute
' API seems to run the program, but you cannot
' see anything.  It's like it's running but
' somewhere way behind.  Here we run PS, but
' then wait for it to exit before we going on.

tStartupInfo.cb = Len(tStartupInfo)
' Here we run PS in a new process.
CreateProcess vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, tStartupInfo, tProcessInfo
' Now, wait from PS to exit.
WaitForSingleObject tProcessInfo.hProcess, INFINITE
' Close the now useless handle.
CloseHandle tProcessInfo.hProcess

Exit Sub

CANNOT_FIND_PROGRAM:
MsgBox "Unable to locate the program.  Please reinstall Picture Scroller.", vbCritical, "Cannot File Program"

End Sub
Private Function IsCommand(ByVal sRunCommand As String, ByVal sLetter As String) As Byte

' Purpose: Search the string given (the first two
'   characters of the command line) for a given letter
'   (the command from Windows).

Dim iPosition As Integer

iPosition = InStr(sRunCommand, sLetter)

If iPosition <> 0 Then IsCommand = True

End Function
