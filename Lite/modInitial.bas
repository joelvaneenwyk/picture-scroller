Attribute VB_Name = "modInitial"
Option Explicit

Public Const RM_NORMAL = 0
Public Const RM_SCROLL = 1

' Holds the program's path; includes a slash at the end.
Public sAppPath As String

' Tells what mode PS is being run in.  (See constants
' above for details.)
Public bRunMode As Byte
Private Sub Main()

' Purpose: Handles what to do when the program first
'   opens.  It also loads previous settings.

On Error GoTo ErrorHandler

Dim sCommandLine As String

' Retrieve the command line parameters.
sCommandLine = UCase(Trim(Command))

' Don't run two Picture Scrollers in normal mode.
CheckPrevInstance sCommandLine

Select Case sCommandLine
    Case ""
        ' Run PS in normal mode.
        bRunMode = RM_NORMAL
    Case Else
        WaitProcess "Loading Picture List", True

        ' Remove unnecessary quotes.
        If Left$(sCommandLine, 1) = """" Then sCommandLine = Right$(sCommandLine, Len(sCommandLine) - 1)
        If Right$(sCommandLine, 1) = """" Then sCommandLine = Left$(sCommandLine, Len(sCommandLine) - 1)

        ' Load the pictures from the file on the
        ' command line.
        LoadSavedList sCommandLine, False, True

        ' If the user pressed cancel, then just stop.
        If bCancelOp = True Then
            EndWaitProcess
            Exit Sub
        Else: EndWaitProcess
        End If

        If tPictureFiles.Count = 0 Then
            MsgBox "This picture list is not valid or does not contain any pictures.", vbInformation
            bRunMode = RM_NORMAL
        Else
            ' Run PS in Normal Mode, scrolling
            ' through the pictures.
            bRunMode = RM_SCROLL

            Load frmScroller
            DoEvents

            Unload frmScroller
            DoEvents
        End If
End Select

' If we're to run PS in normal mode or screen
' saver config mode, then load the main form.
If bRunMode = RM_NORMAL Then
    Load frmMain
    DoEvents
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub CheckPrevInstance(ByVal sCommandLine As String)

' Purpose: Make sure another instance of PS
'   isn't running.  Bring up the other if one
'   does exist and unload us; otherwise we can
'   load the settings.

On Error GoTo ErrorHandler

Dim sTitle As String

' See if another instance of PS is open.
' Note: We don't care if another PS is open, if we
'   are to run in ScreenSaver mode.
If App.PrevInstance = True And sCommandLine <> "/SAVER_RUN" Then
    sTitle = App.Title
    App.Title = Hex$(App.hInstance)
    AppActivate sTitle
    SendKeys "% R", True

    End
Else
    ' If there isn't another PS then we're safe
    ' and can now load the settings.

    ' Figure out the program's path
    sAppPath = NormalizePath(App.Path)

    ' Retrieve the user options
    Open sAppPath & "Options.dat" For Binary As FILENUM_OPTIONS
    Get #1, , tProgramOptions
    Close #1

    ' Detect whether this is Windows NT or not.
    bIsWinNT = IsWinNT
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
