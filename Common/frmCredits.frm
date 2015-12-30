VERSION 5.00
Begin VB.Form frmCredits 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrPrint 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   105
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PHOTOWIDTH As Integer = 164
Const GAP As Integer = 40

Dim iStartX As Integer
Dim iStartY As Integer
Dim Index As Integer
Dim iOffset As Integer
Dim bDraw As Byte
Dim bGroup As Byte
Dim iLargest As Integer

Dim nHeight As Integer

Dim iX2 As Integer
Dim iY2 As Integer

Dim iPos() As Integer
Dim sLines() As String

Dim dixuPhotos() As DirectDrawSurface7
Dim dixuBackground As DirectDrawSurface7

Dim bUnload As Byte
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim iDisplay As Integer
Dim tStore As DISPLAY_MODE
Dim bScreenSetting As Byte

For nIndex = 1 To DirectDraw.ModeCount
    If DirectDraw.ModeWidth(nIndex) = 640 And DirectDraw.ModeHeight(nIndex) = 480 And DirectDraw.ModeBPP(nIndex) = 16 Then
        iDisplay = nIndex
        Exit For
    End If
Next nIndex

bScreenSetting = tProgramOptions.bScreenSetting

If iDisplay = 0 Then
    MsgBox "Your computer does not seem to support" & vbCr & "the optimum resolution for viewing the credits.", vbInformation
    tProgramOptions.bScreenSetting = 0
Else
    ' Store the user's settings.
    tStore = tProgramOptions.tDisplayMode

    ' Set tProgramOptions to the resolution we want
    ' temporarally.
    With tProgramOptions.tDisplayMode
        .iWidth = DirectDraw.ModeWidth(iDisplay)
        .iHeight = DirectDraw.ModeHeight(iDisplay)
        .bBPP = DirectDraw.ModeBPP(iDisplay)
    End With

    tProgramOptions.bScreenSetting = 3
End If

If DirectDraw.InitDirectDraw(Me.hwnd, 0) = False Then
    DirectDraw.KillDirectDraw

    GoSub RestoreSettings

    MsgBox "An error occurred while trying to initialize DirectX.  Please be sure that you have DirectX 7.0 or greater installed on your computer and that no other program is currently using DirectX."

    Exit Sub
End If

SetMouseCount False

dixuBackBuffer.SetFont Me.Font

MainLoop

SetMouseCount True

Set dixuBackBuffer = Nothing

DirectDraw.KillDirectDraw

GoSub RestoreSettings

Exit Sub

RestoreSettings:
tProgramOptions.bScreenSetting = bScreenSetting
If iDisplay <> 0 Then
    tProgramOptions.tDisplayMode = tStore
End If

Return

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

bUnload = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

tmrPrint.Enabled = False
bUnload = True

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

bUnload = False

End Sub
Private Sub MainLoop()

On Error GoTo ErrorHandler

Dim nIndexGroups As Integer
Dim nIndexLines As Integer
Dim nIndex As Integer

Dim iPosX As Integer
Dim iPosY As Integer

Dim iWidth As Integer
Dim iHeight As Integer

Dim lLastTime As Long

Dim dixuTemp As DirectDrawSurface7

bGroup = LoadResString(999)

ReDim sLines(1 To LoadResString(998) * bGroup)
ReDim iPos(1 To UBound(sLines))
ReDim dixuPhotos(1 To LoadResString(998))

For nIndexGroups = 0 To LoadResString(998) - 1
    For nIndexLines = 1 To bGroup
        sLines((nIndexGroups * bGroup) + nIndexLines) = LoadResString(999 + nIndexLines + nIndexGroups * 10)
    Next nIndexLines

    DirectDraw.CreateSurface dixuPhotos(nIndexGroups + 1), 0, 0, CREATE_FROM_RES, 20 + nIndexGroups, "PICS"
Next nIndexGroups

nHeight = Me.TextHeight("A")

DirectDraw.CreateSurface dixuTemp, 0, 0, CREATE_FROM_RES, 30, "PICS"
DirectDraw.CreateSurface dixuBackground, iScreenWidth, iScreenHeight, CREATE_FROM_NONE, 0

Do
    If iScreenWidth - iPosX < 128 Then
        iWidth = iScreenWidth - iPosX
    Else: iWidth = 128
    End If

    If iScreenHeight - iPosY < 128 Then
        iHeight = iScreenHeight - iPosY
    Else: iHeight = 128
    End If

    DirectDraw.BltFast dixuBackground, iPosX, iPosY, iWidth, iHeight, dixuTemp, 0, 0, False

    If iPosX + 128 < iScreenWidth Then
        iPosX = iPosX + 128
    Else
        iPosX = 0
        iPosY = iPosY + 128
    End If

    If iPosY >= iScreenHeight Then Exit Do
Loop

Set dixuTemp = Nothing

iStartX = DirectDraw.Width(dixuPhotos(1)) + GAP * 2
iStartY = (iScreenHeight - ((bGroup * nHeight) + ((bGroup - 1) * 40))) / 2

iLargest = iScreenWidth - GAP * 2 - iStartX

iX2 = iStartX + iLargest + GAP
iY2 = iStartY + (bGroup * Me.TextHeight("A")) + ((bGroup - 1) * 40)

lLastTime = DirectX.TickCount

Index = 1

bDraw = 1

dixuBackBuffer.SetForeColor &HFFFF&
dixuBackBuffer.SetFillColor 0
dixuBackBuffer.SetFillStyle 0

Do
    If DirectDraw.TestState = False Then
        ShowCursor True

        Do Until DirectDraw.TestState
            DoEvents
        Loop

        ShowCursor False

        ' Clear the back buffer of any garbage.
        DirectDraw.ClearSurface dixuBackBuffer

        If tmrPrint.Tag <> "" Then
            tmrPrint.Enabled = tmrPrint.Tag
            tmrPrint.Tag = ""
        End If
    End If

    If DirectX.TickCount - lLastTime > 30 And bDraw = 1 Then
        For nIndex = 1 To Index
            If iPos(nIndex + iOffset) = 0 Then iPos(nIndex + iOffset) = 1

            If nIndex = Index And iPos(nIndex + iOffset) = Len(sLines(nIndex + iOffset)) Then
                If Index = 6 Then
                    bDraw = 0
                    tmrPrint.Enabled = True
                    Exit For
                Else: Index = Index + 1
                End If
            Else: iPos(nIndex + iOffset) = iPos(nIndex + iOffset) + 1
            End If
        Next nIndex

        lLastTime = DirectX.TickCount
    End If

    DirectDraw.BltFast dixuBackBuffer, 0, 0, iScreenWidth, iScreenHeight, dixuBackground, 0, 0, False

    dixuBackBuffer.DrawBox iStartX, iStartY, iX2, iY2

    For nIndex = 1 To Index
        dixuBackBuffer.DrawText iStartX + GAP / 2, iStartY + (40 * nIndex), Left$(sLines(nIndex + iOffset), iPos(nIndex + iOffset)), False
    Next nIndex

    DirectDraw.BltFast dixuBackBuffer, GAP, (iScreenHeight - DirectDraw.Height(dixuPhotos(iOffset \ 6 + 1))) \ 2, 0, 0, dixuPhotos(iOffset \ 6 + 1), 0, 0, False

    If bUnload = True Then Exit Do

    DirectDraw.Flip

    DirectDraw.ClearSurface dixuBackBuffer

    DoEvents
Loop

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrPrint_Timer()

On Error GoTo ErrorHandler

Dim nIndex As Integer

Index = 1

If iOffset = UBound(sLines) - bGroup Then
    ReDim iPos(1 To bGroup * LoadResString(998))
    iOffset = 0
Else
    iLargest = 0

    Index = 1

    iOffset = iOffset + bGroup
End If

iStartX = DirectDraw.Width(dixuPhotos(iOffset \ bGroup + 1)) + GAP * 2

iLargest = iScreenWidth - GAP * 2 - iStartX

iX2 = iStartX + iLargest + GAP
iY2 = iStartY + (bGroup * nHeight) + ((bGroup - 1) * 40)

bDraw = 1

tmrPrint.Enabled = False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
