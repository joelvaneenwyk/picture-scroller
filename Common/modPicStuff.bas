Attribute VB_Name = "modPicStuff"
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

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As CBoolean, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As CBoolean, riid As GUID, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Function ConfirmType(ByVal sFileName As String) As Byte

' Purpose: See if the file is any of the supported
'   types if it isn't then the function will
'   return nothing.

On Error GoTo ErrorHandler

Dim iPosition As Integer
Dim sExtension As String

If FileExists(sFileName) = False Then Exit Function

For iPosition = Len(sFileName) To 1 Step -1
    If Mid$(sFileName, iPosition, 1) = "." Then Exit For
Next iPosition

' Extract the extension from the file name
sExtension = LCase(Right$(sFileName, Len(sFileName) - iPosition))

' We put a different number for each picture so that
' the correct icon will for shown for each type
If sExtension = "bmp" Then
    ConfirmType = 1
ElseIf sExtension = "jpg" Or sExtension = "jpeg" Or sExtension = "jpe" Or sExtension = "jfif" Then
    ConfirmType = 2
ElseIf sExtension = "gif" Then
    ConfirmType = 3
ElseIf sExtension = "wmf" Then
    ConfirmType = 4
ElseIf sExtension = "emf" Then
    ConfirmType = 5
ElseIf sExtension = "ico" Then
    ConfirmType = 6
ElseIf IsPCX(sFileName) = True Then
    ConfirmType = PIC_PCX
ElseIf IsPSD(sFileName) = True Then
    ConfirmType = PIC_PSD
ElseIf IsTGA(sFileName) = True Then
    ConfirmType = PIC_TGA
ElseIf IsLBM(sFileName) = True Then
    ConfirmType = PIC_LBM
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function LoadSavedList(ByVal sFileName As String, ByVal bShowMessage As Byte, ByVal bSetRecent As Byte) As Byte

' Purpose: Loads a list of picture files from a
'   previously saved list and adds them to the collection.
'   Returns FALSE if the file doesn't exist or so corrupt.

On Error GoTo ErrorHandler

Dim sLine As String
Dim bType As Byte
Dim iLastSlash As Integer
Dim lFolderID As Long
Dim bGoOn As Byte

If FileExists(sFileName) = False Then
    ' We cannot load a file that isn't there!
    LoadSavedList = False
    Exit Function
End If

Open sFileName For Input As FILENUM_LIST

If EOF(FILENUM_LIST) Then
    Close FILENUM_LIST

    LoadSavedList = True
    Exit Function
End If

' Retrieve the first line of the file.
Line Input #FILENUM_LIST, sLine

If sLine <> LIST_HEADER Then
    Close FILENUM_LIST

    Exit Function
End If

Do Until EOF(FILENUM_LIST) Or bCancelOp = True
    Line Input #FILENUM_LIST, sLine

    bType = ConfirmType(sLine)

    If bType <> 0 Then
        iLastSlash = LastSlash(sLine)

        If iLastSlash <> 0 Then
            lFolderID = GetNewFolderID(Left$(sLine, iLastSlash))

            tPictureFiles.Add ReturnCPicture(sLine, bType, lFolderID)
        End If
    ElseIf bShowMessage = True And bGoOn = False Then
        If MsgBox("This list of pictures contains references to files that cannot be accessed right now.  Are you sure you want to continue?", vbYesNo) = vbNo Then
            ' The user choose to cancel, so no
            ' more messages.
            LoadSavedList = True

            Close FILENUM_LIST

            bCancelOp = True
            Exit Function
        Else: bGoOn = True
        End If
    End If

    DoEvents
Loop

Close FILENUM_LIST

If bSetRecent = True Then
    ' Store the file name as a recently opened file.
    With tProgramOptions
        ' We don't want both slots to say the same thing.
        If .sRecent1 <> sFileName Then
            .sRecent2 = .sRecent1
            .sRecent1 = sFileName
        End If
    End With
End If

' We have successfully added to the pictures.
LoadSavedList = True

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function GetNewFolderID(ByVal sFolder As String) As Long

' Purpose: Get the folder ID of this folder if it
'   hasn't been added before, add it to the collection

On Error GoTo ErrorHandler

Dim lFolderID As Long

lFolderID = GetFolderID(sFolder)

If lFolderID = 0 Then
    lFolderID = tAddedFolders.Count + 1

    tAddedFolders.Add ReturnCFolder(sFolder, lFolderID), "DIR_" & lFolderID
End If

GetNewFolderID = lFolderID

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function GetFolderID(ByVal sFolder As String) As Long

' Purpose: Gets the folder ID of the given folder

On Error GoTo ErrorHandler

Dim nIndex As Long

For nIndex = 1 To tAddedFolders.Count
    If tAddedFolders(nIndex).FolderName = sFolder Then
        GetFolderID = tAddedFolders(nIndex).FolderID
        Exit For
    End If
Next nIndex

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function ReturnCPicture(ByVal sFileName As String, ByVal bType As Byte, ByVal lFolderID As Long) As CPictureInfo

' Purpose: Takes the correct parameters and returns
'   a CPictureInfo class

On Error GoTo ErrorHandler

Dim objPicInfo As New CPictureInfo

With objPicInfo
    .FileName = sFileName
    .PicType = bType
    .FolderID = lFolderID
End With

Set ReturnCPicture = objPicInfo

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function ReturnCFolder(ByVal sFolder As String, ByVal lFolderID As Long) As CFolderInfo

' Purpose: Takes the correct parameters and returns
'   a CPictureInfo class

On Error GoTo ErrorHandler

Dim objFolder As New CFolderInfo

With objFolder
    .FolderName = sFolder
    .FolderID = lFolderID
End With

Set ReturnCFolder = objFolder

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Public Function GetPicFromIndex(picDest As PictureBox, ByVal lPictureIndex As Long, Optional iWidth As Integer, Optional iHeight As Integer, Optional bProportional As Byte) As Byte

' Purpose: Load a picture (of any supported type)
'   into the given PictureBox.  picWork and picDest
'   can (and should be) the same PictureBox.

On Error GoTo LoadError

Dim objPicture As StdPicture
Dim hPictureDC As Long

Dim pImage As IMAGEFILE
Dim tBitmap As BITMAPINFO
Dim tBitmapData As Long

Dim bStretchIt As Byte

Dim iNewWidth As Integer
Dim iNewHeight As Integer

picDest.Picture = LoadPicture()

' Load the picture
Select Case tPictureFiles(lPictureIndex).PicType
    Case Is <= 6
        ' If we are to load the picture in specific
        ' dimensions, then first put the picture
        ' in our own objPicture.
        If iWidth <> 0 And iHeight <> 0 And tPictureFiles(lPictureIndex).PicType <> 4 And tPictureFiles(lPictureIndex).PicType <> 5 Then
            bStretchIt = True

            Set objPicture = LoadPicture(tPictureFiles(lPictureIndex).FileName)

            ' Create a DC for the picture
            hPictureDC = CreateCompatibleDC(0)
            ' Combine picture and DC
            SelectObject hPictureDC, objPicture.Handle
        Else
            picDest.Picture = LoadPicture(tPictureFiles(lPictureIndex).FileName)

            If iWidth <> 0 And iHeight <> 0 Then
                picDest.Width = iWidth
                picDest.Height = iHeight
            End If
        End If
    Case PIC_PCX
        LoadPCX tPictureFiles(lPictureIndex).FileName, pImage
    Case PIC_PSD
        LoadPSD tPictureFiles(lPictureIndex).FileName, pImage
    Case PIC_TGA
        LoadTGA tPictureFiles(lPictureIndex).FileName, pImage
    Case PIC_LBM
        LoadLBM tPictureFiles(lPictureIndex).FileName, pImage
End Select

' If the picture is a "special" type, then we need
' to put it into objPictureBox; if it was a "regular"
' type, then stretch it if we are specified dimensions.
If tPictureFiles(lPictureIndex).PicType > 6 Then
    If iWidth = 0 Then iWidth = pImage.Width
    If iHeight = 0 Then iHeight = pImage.Height

    picDest.Width = iWidth
    picDest.Height = iHeight

    DrawImage picDest.hdc, pImage, iWidth, iHeight
    picDest.Picture = picDest.Image
ElseIf bStretchIt = True Then
    With picDest
        If bProportional = True Then
            iNewWidth = iWidth
            iNewHeight = iHeight

            GetProportional DirectDraw.HimetricToPixel(objPicture.Width, Screen.TwipsPerPixelX), DirectDraw.HimetricToPixel(objPicture.Height, Screen.TwipsPerPixelY), iNewWidth, iNewHeight

            .Width = iNewWidth
            .Height = iNewHeight
        Else
            .Width = iWidth
            .Height = iHeight
        End If

        SetStretchBltMode .hdc, STRETCH_DELETESCANS
        StretchBlt .hdc, 0, 0, .Width, .Height, hPictureDC, 0, 0, DirectDraw.HimetricToPixel(objPicture.Width, Screen.TwipsPerPixelX), DirectDraw.HimetricToPixel(objPicture.Height, Screen.TwipsPerPixelY), vbSrcCopy
    End With

    DeleteDC hPictureDC
    Set objPicture = Nothing
End If

GetPicFromIndex = True

Exit Function

LoadError:
ErrHandle
On Error GoTo 0
GetPicFromIndex = False

End Function
Public Function PictureFromBits(abPic() As Byte) As IPicture  ' not a StdPicture!!

' Purpose: Loads a picture from a string of bytes.  Used
'   internally to load a picture from the resource file.

On Error GoTo ErrorHandler

Dim nLow As Long
Dim cbMem  As Long
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown
Dim ipic As IPicture

' Get the size of the picture's bits
nLow = LBound(abPic)
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

Exit Function

ErrorHandler:
ErrHandle
On Error GoTo 0

End Function
