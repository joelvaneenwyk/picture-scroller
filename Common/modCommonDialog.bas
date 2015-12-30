Attribute VB_Name = "modCommonDialog"
Option Explicit

Public Enum FILTER_TYPES
    MUSIC = 0
    PICTURES = 1
    SAVED_LIST = 2
End Enum

' Error constants when running files
Const ERROR_FILE_NOT_FOUND = 2&
Const ERROR_PATH_NOT_FOUND = 3&
Const ERROR_GEN_FAILURE = 31&

' Used for file dialog
Const OFN_EXPLORER = &H80000
Const OFN_LONGNAMES = &H200000
Const OFN_HIDEREADONLY = &H4
Const OFN_PATHMUSTEXIST = &H800
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_ALLOWMULTISELECT = &H200
Const OFN_OVERWRITEPROMPT = &H2
Const FNERR_BUFFERTOOSMALL = &H3003

' Used for color dialog
Const CC_RGBINIT = &H1
Const CC_FULLOPEN = &H2
Const CC_ANYCOLOR = &H100

' Type for holding color information for Color dialog
Private Type ChooseColor
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

' Type for holding open and save file dialog information
Private Type OpenFileName
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' Functions for showing the open/save common dialog box
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Function OpenFileDialog(ByVal hwnd As Long, ByVal bFilter As FILTER_TYPES, sSelectedFiles() As String) As String

' Purpose: Show the Open File CommonDialog window and
'   returns the file(s) selected

On Error GoTo ErrorHandler

Dim tDialog As OpenFileName

Dim lReturnVal As Long
Dim sPath As String
Dim iPosition As Integer
Dim iNextPosition As Integer

ReDim sSelectedFiles(0)

' Prepare to set dialog title property and show the save
' dialog box
With tDialog
    .lStructSize = Len(tDialog)
    .hwndOwner = hwnd
    .hInstance = App.hInstance

    Select Case bFilter
        Case PICTURES
            .lpstrFilter = "All Files (*.*)" & vbNullChar & _
                "*.*" & vbNullChar & "All Picture Files" & _
                vbNullChar & "*.bmp;*.jpg;*.gif;*.wmf;" & _
                "*.emf;*.ico;*.pcx;*.psd;*.tga;*.lbm" & _
                vbNullChar & "Windows Bitmaps (*.bmp)" & _
                vbNullChar & "*.bmp" & vbNullChar & _
                "JPEG Files (*.jpg)" & vbNullChar & "*.jpg" & _
                vbNullChar & "GIF Files (*.gif)" & _
                vbNullChar & "*.gif" & vbNullChar & _
                "Metafiles (*.wmf, *.emf)" & vbNullChar & _
                "*.wmf;*.emf" & vbNullChar & "Icons (*.ico)" & _
                vbNullChar & "*.ico" & vbNullChar & _
                "PCX Files (*.pcx)" & vbNullChar & "*.pcx" & _
                vbNullChar & "PhotoShop Files (*.psd)" & _
                vbNullChar & "*.psd" & vbNullChar & _
                "TGA Files (*.tga)" & vbNullChar & "*.tga" & _
                vbNullChar & "DPaint Files (*.lbm)" & vbNullChar & "*.lbm" & vbNullChar

            .lpstrTitle = "Add Picture File(s)"
            .lpstrInitialDir = tProgramOptions.sLastPicFolder
            .flags = OFN_ALLOWMULTISELECT
        Case MUSIC
            .lpstrFilter = "All Files (*.*)" & vbNullChar & _
                "*.*" & vbNullChar & "All Music Files" & _
                vbNullChar & "*.wav;*.mid" & vbNullChar & _
                "Wave Files (*.wav)" & vbNullChar & "*.wav" & _
                vbNullChar & "MIDI Files (*.mid)" & _
                vbNullChar & "*.mid" & vbNullChar

            .lpstrTitle = "Select Background Music"
            .lpstrInitialDir = tProgramOptions.sLastMusicFolder
        Case SAVED_LIST
            .lpstrFilter = "Saved List (*.pcs)" & vbNullChar & "*.pcs" & vbNullChar

            .lpstrTitle = "Open Saved List"
            .lpstrInitialDir = tProgramOptions.sLastSavedFolder
    End Select

    .flags = .flags Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY

    ' If no initial directory has been set, use the
    ' application path (that path contains the
    ' example pcs file).
    If .lpstrInitialDir = "" Then .lpstrInitialDir = sAppPath

    ' Select the default filter to use.
    If bFilter = MUSIC Or PICTURES Then
        .nFilterIndex = 2
    Else: .nFilterIndex = 1
    End If
End With

GetFileName:
With tDialog
    .lpstrFile = String(tProgramOptions.MAX_LEN, 0)
    .nMaxFile = tProgramOptions.MAX_LEN
End With

lReturnVal = GetOpenFileName(tDialog)

' This is used to find any errors that may have occurred
' during the retrieval of the file name
If lReturnVal = 0 Then
    lReturnVal = CommDlgExtendedError
    If lReturnVal > 0 Then
        If lReturnVal = FNERR_BUFFERTOOSMALL Then
            tProgramOptions.MAX_LEN = Asc(Left$(tDialog.lpstrFile, 1)) + Asc(Mid(tDialog.lpstrFile, 2, 1))

            MsgBox "The file(s) you selected surpassed the excepted maximum length (greater than " & tDialog.nMaxFile & " characters!).  However, this has been corrected, so please try again.", vbCritical, "Maximum Length Exceeded"
            GoTo GetFileName
        Else
            MsgBox "Sorry, but an unexpected error occurred.  Please try again.", vbCritical, "Error!"
            Exit Function
        End If
    End If
End If

With tDialog
    ' Make sure the user selected a file
    If .nFileOffset = 0 Then Exit Function

    ' Extract the path from the returned string
    sPath = Left$(.lpstrFile, .nFileOffset - 1)
    NormalizePath sPath

    ' Start searching after the path
    iPosition = .nFileOffset

    Do
        ' Find the next division before the next file
        iNextPosition = InStr(iPosition + 1, .lpstrFile, vbNullChar)

        ' If two null characters are together
        ' then we've got all the files
        If iNextPosition - iPosition = 1 Then Exit Do

        ' Redim the array if necessary
        If sSelectedFiles(0) <> "" Then ReDim Preserve sSelectedFiles(UBound(sSelectedFiles) + 1)

        ' Put the file name with the path into the array
        sSelectedFiles(UBound(sSelectedFiles)) = sPath & Mid(.lpstrFile, iPosition + 1, iNextPosition - iPosition - 1)

        ' Start searching the next time after the
        ' end of the file just found
        iPosition = iNextPosition
    Loop

    ' Save the directory just used for next time.
    Select Case bFilter
        Case PICTURES
            tProgramOptions.sLastPicFolder = sPath
        Case MUSIC
            tProgramOptions.sLastMusicFolder = sPath
        Case SAVED_LIST
            tProgramOptions.sLastSavedFolder = sPath
    End Select

    ' Return the path that was selected
    OpenFileDialog = sPath
End With

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function SaveFileDialog(ByVal hwnd As Long) As String

' Purpose: Show the Save File CommonDialog window.

On Error GoTo ErrorHandler

Dim tDialog As OpenFileName

Dim lReturnVal As Long

' Prepare to set dialog title property and show the save
' dialog box
With tDialog
    .lStructSize = Len(tDialog)
    .hwndOwner = hwnd
    .hInstance = App.hInstance
    .lpstrTitle = "Save Current List"
    .lpstrFilter = "Saved List (*.pcs)" & vbNullChar & "*.pcs" & vbNullChar
    .lpstrDefExt = "pcs"
    .lpstrInitialDir = tProgramOptions.sLastSavedFolder
    .flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT

GetFileName:
    .lpstrFile = String(tProgramOptions.MAX_LEN, 0)
    .nMaxFile = tProgramOptions.MAX_LEN
End With

lReturnVal = GetSaveFileName(tDialog)

' This is used to find any errors that may have occurred
' during the retrieval of the file name
If lReturnVal = 0 Then
    lReturnVal = CommDlgExtendedError
    If lReturnVal > 0 Then
        If lReturnVal = FNERR_BUFFERTOOSMALL Then
            tProgramOptions.MAX_LEN = Asc(Left$(tDialog.lpstrFile, 1)) + Asc(Mid(tDialog.lpstrFile, 2, 1))

            MsgBox "The file name you entered is longer than the excepted maximum length (greater than " & tDialog.nMaxFile & " characters!).  However, this has been corrected, so please try again.", vbCritical, "Maximum Length Exceeded"
            GoTo GetFileName
        Else
            MsgBox "Sorry, but an unexpected error occurred.  Please try again.", vbCritical, "Error!"
            Exit Function
        End If
    End If
End If

If tDialog.nFileOffset <> 0 Then
    tProgramOptions.sLastSavedFolder = Left$(tDialog.lpstrFile, tDialog.nFileOffset)

    SaveFileDialog = TrimNulls(tDialog.lpstrFile)
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function SelectColor(ByVal hwnd As Long, ByVal lInitialColor As Long) As Long

On Error GoTo ErrorHandler

Dim tColorDialog As ChooseColor
Dim lReturnVal As Long

With tColorDialog
    .lStructSize = Len(tColorDialog)
    .hwndOwner = hwnd
    .hInstance = App.hInstance
    .lpCustColors = StrConv(tProgramOptions.bCustomColors, vbUnicode)
    .rgbResult = lInitialColor
    .flags = CC_RGBINIT Or CC_FULLOPEN Or CC_ANYCOLOR
End With

lReturnVal = ChooseColor(tColorDialog)

If lReturnVal <> 0 Then
    ' Store any custom colors the user created
    tProgramOptions.bCustomColors() = StrConv(tColorDialog.lpCustColors, vbFromUnicode)

    ' Return the color selected
    SelectColor = tColorDialog.rgbResult
Else: SelectColor = -1
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
