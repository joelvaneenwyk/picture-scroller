VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Pictures"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1260
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.Slider sldOrder 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      LargeChange     =   1
      Min             =   1
      Max             =   2
      SelStart        =   1
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.PictureBox picShown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   735
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   345
      Left            =   105
      TabIndex        =   9
      Top             =   4650
      Width           =   1515
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1515
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "&Setup"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1515
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1515
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Sele&ct Picture(s)"
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1515
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   5100
      Left            =   1785
      ScaleHeight     =   336
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   336
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   5100
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   105
         ScaleHeight     =   319
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   105
         Width           =   3135
      End
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   4305
      Width           =   1515
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      Caption         =   "Front"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      Caption         =   "Back"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblOrder 
      AutoSize        =   -1  'True
      Caption         =   "Picture Stacking:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   7
      X2              =   109
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Menu mnuPicOptions 
      Caption         =   "&PicOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuOptRestore 
         Caption         =   "&Restore"
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bNoPrint As Byte

Dim fPageScale As Single
Dim bOver As Byte
Dim bSelected As Byte
Dim bDown As Byte
Dim pPos As POINTAPI
Dim bByCode As Byte
Dim sEdge As String
Dim iValueX As Integer
Dim iValueY As Integer

Private Const BF_RECT = &HF
Private Const EDGE_BUMP = &H9

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal EDGE As Long, ByVal grfFlags As Long) As Long
Private Sub cmdAdd_Click()

On Error GoTo ErrorHandler

Dim nIndex As Byte

' Allow the user to change their selection.
frmInsertPic.Show 1
DoEvents

' If they've changed something...
If frmInsertPic.bChanged = True Then
    WaitProcess "Loading pictures", True

    ' Remove all the previously loaded pictures.
    For nIndex = 1 To picShown.UBound
        Unload picShown(nIndex)
    Next nIndex

    ' Load the pictures into PictureBoxes
    For nIndex = 1 To UBound(tSelectedPics)
        DoEvents

        Load picShown(nIndex)

        GetPicFromIndex picHidden, tSelectedPics(nIndex).lPictureIndex

        With tSelectedPics(nIndex).tActualSize
            .Right = picHidden.Width * fPageScale
            .Bottom = picHidden.Height * fPageScale

            picShown(nIndex).Width = .Right
            picShown(nIndex).Height = .Bottom

            picShown(nIndex).Picture = LoadPicture()

            SetStretchBltMode picShown(nIndex).hdc, STRETCH_DELETESCANS
            StretchBlt picShown(nIndex).hdc, 0, 0, .Right, .Bottom, picHidden.hdc, 0, 0, picHidden.Width, picHidden.Height, vbSrcCopy

            picShown(nIndex).Picture = picShown(nIndex).Image

            picHidden.Picture = LoadPicture()
        End With
    Next nIndex

    If picShown.UBound > 1 Then
        sldOrder.Max = picShown.UBound
        sldOrder.Enabled = True
    Else
        sldOrder.Max = 2
        sldOrder.Enabled = False
    End If

    DoEvents

    DrawPage

    EndWaitProcess
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdHelp_Click()

ShowHelp Me.hwnd, "Options_Preview.htm", True

End Sub
Private Sub cmdPrint_Click()

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim rRect As RECT

If UBound(tSelectedPics) = 0 Then
    MsgBox "Please select the pictures you would like to print by clicking the 'Select Picture(s)' button.", vbInformation
    Exit Sub
ElseIf MsgBox("Press 'OK' to begin printing.", vbInformation + vbOKCancel) = vbCancel Then
    Exit Sub
End If

WaitProcess "Printing", True, Me

For nIndex = picShown.UBound To 1 Step -1
    With tSelectedPics(nIndex).tActualSize
        If tSelectedPics(nIndex).iValueX < 0 Then
            rRect.Left = .Left + .Right - 2
        Else: rRect.Left = .Left
        End If

        If tSelectedPics(nIndex).iValueY < 0 Then
            rRect.Top = .Top + .Bottom
        Else: rRect.Top = .Top
        End If

        rRect.Left = (rRect.Left / fPageScale)
        rRect.Top = (rRect.Top / fPageScale)
        rRect.Right = .Right * tSelectedPics(nIndex).iValueX / fPageScale
        rRect.Bottom = .Bottom * tSelectedPics(nIndex).iValueY / fPageScale

        DoEvents

        GetPicFromIndex picHidden, tSelectedPics(nIndex).lPictureIndex

        DoEvents

        With rRect
            Printer.PaintPicture picHidden.Picture, .Left, .Top, .Right, .Bottom
        End With

        DoEvents

        picHidden.Picture = LoadPicture()

        If bCancelOp = True Then
            Printer.KillDoc
            GoTo Done
        Else: DoEvents
        End If
    End With
Next nIndex

Printer.EndDoc

Done:
EndWaitProcess Me

Exit Sub

ErrorHandler:
On Error GoTo 0
MsgBox "Unable to print to the printer.  Please ensure that it is ready.", vbExclamation
ErrHandle

End Sub
Private Sub cmdRemove_Click()

On Error GoTo ErrorHandler

Dim nIndex As Byte

If UBound(tSelectedPics) = 0 Then
    MsgBox "Click the 'Add' button to add the pictures you would like to print.  If you want to remove one later, click this button.", vbInformation
    Exit Sub
ElseIf bSelected = 0 Then
    MsgBox "Select the picture you would like to remove from those printed, then click this button.", vbInformation
    Exit Sub
End If

For nIndex = bSelected To picShown.UBound - 1
    picShown(nIndex) = picShown(nIndex + 1)
    tSelectedPics(nIndex) = tSelectedPics(nIndex + 1)
Next nIndex

Unload picShown(picShown.UBound)

If UBound(tSelectedPics) = 1 Then
    ReDim tSelectedPics(0)
Else: ReDim Preserve tSelectedPics(1 To UBound(tSelectedPics) - 1)
End If

If picShown.UBound > 1 Then
    sldOrder.Max = picShown.UBound
    sldOrder.Enabled = True
Else
    sldOrder.Max = 2
    sldOrder.Enabled = False
End If

DrawPage

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdSetup_Click()

On Error GoTo ErrorHandler

Screen.MousePointer = 11

ShowPrinter Me, PD_HIDEPRINTTOFILE Or PD_NOSELECTION Or PD_NOWARNING Or PD_NOPAGENUMS

Screen.MousePointer = 0

SizePage

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
    Unload Me
End If

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim nIndex As Long

If Printers.Count = 0 Then
    bNoPrint = True

    cmdSetup.Enabled = False
    cmdPrint.Enabled = False
Else
    bNoPrint = False

    cmdSetup.Enabled = True
    cmdPrint.Enabled = True

    Printer.ScaleMode = vbPixels
End If

SizePage

ReDim tSelectedPics(0)

If bNoPrint = True Then
    MsgBox "Since there are no printers installed on this computer, you cannot actually print," & vbCr & "but you can still see how Picture Scroller enables you to print pictures.", vbInformation
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub SizePage()

On Error GoTo ErrorHandler

Dim iWidth As Integer
Dim iHeight As Integer
Dim iNewWidth As Integer
Dim iNewHeight As Integer
Dim sScale As String

With picPage
    If bNoPrint = False Then
        iWidth = Printer.ScaleWidth
        iHeight = Printer.ScaleHeight
    Else
        iWidth = 2000
        iHeight = 3000
    End If

    iNewWidth = picBackground.ScaleWidth - 20
    iNewHeight = picBackground.ScaleHeight - 20

    fPageScale = GetProportional(iWidth, iHeight, iNewWidth, iNewHeight)

    .Width = iNewWidth
    .Height = iNewHeight

    If .Left <> (picBackground.ScaleWidth - iNewWidth) \ 2 Then .Left = (picBackground.ScaleWidth - iNewWidth) \ 2
    If .Top <> (picBackground.ScaleHeight - iNewHeight) \ 2 Then .Top = (picBackground.ScaleHeight - iNewHeight) \ 2

    lblScale.Caption = "Scale: " & Round(fPageScale * 100) & "% of Actual"

    DoEvents
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub DrawPage()

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim rRect As RECT

picPage.Picture = LoadPicture()

For nIndex = picShown.UBound To 1 Step -1
    With tSelectedPics(nIndex).tActualSize
        If tSelectedPics(nIndex).iValueX < 0 Then
            rRect.Left = .Left + .Right - 2
        Else: rRect.Left = .Left
        End If

        If tSelectedPics(nIndex).iValueY < 0 Then
            rRect.Top = .Top + .Bottom
        Else: rRect.Top = .Top
        End If

        SetStretchBltMode picPage.hdc, STRETCH_DELETESCANS
        StretchBlt picPage.hdc, rRect.Left, rRect.Top, .Right * tSelectedPics(nIndex).iValueX, .Bottom * tSelectedPics(nIndex).iValueY, picShown(nIndex).hdc, 0, 0, picShown(nIndex).ScaleWidth, picShown(nIndex).ScaleHeight, vbSrcCopy

        rRect.Left = .Left
        rRect.Top = .Top
        rRect.Right = rRect.Left + .Right
        rRect.Bottom = rRect.Top + .Bottom

        If tSelectedPics(nIndex).iValueX < 0 Then
            rRect.Left = rRect.Left - 1
        End If

        If tSelectedPics(nIndex).iValueY < 0 Then
            rRect.Bottom = rRect.Bottom + 1
        End If

        If nIndex = bSelected Then
            DrawEdge picPage.hdc, rRect, EDGE_BUMP, BF_RECT
        ElseIf nIndex = bOver Then
            DrawFocusRect picPage.hdc, rRect
        End If
    End With
Next nIndex

picPage.Refresh

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrorHandler

frmMain.Enabled = True

frmMain.Show
DoEvents

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub mnuOptRestore_Click()

On Error GoTo ErrorHandler

With tSelectedPics(bSelected)
    .tActualSize.Right = picShown(bSelected).Width
    .tActualSize.Bottom = picShown(bSelected).Height
    .iValueX = 1
    .iValueY = 1
End With

DrawPage

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picPage_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

If KeyCode = vbKeyDelete Then cmdRemove_Click

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If Button = 1 And bOver >= 1 Then
    bSelected = bOver

    DrawPage

    With tSelectedPics(bSelected).tActualSize
        pPos.X = X - .Left
        pPos.Y = Y - .Top
    End With

    If picShown.UBound > 1 Then
        bByCode = True
        If sldOrder.Enabled = False Then sldOrder.Enabled = True
        sldOrder.Value = sldOrder.Max - (bSelected - 1)
        bByCode = False
    End If

    bDown = 1
ElseIf Button = 2 And bOver <> 0 Then
    Me.PopupMenu mnuPicOptions, , picBackground.Left + picPage.Left + X, picBackground.Top + picPage.Top + Y
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim bOverPic As Byte
Dim bTemp As Byte
Dim iTempUno As Integer
Dim iTempDos As Integer
Dim rRect As RECT

If bDown = 1 Then
    If sEdge = "" Then
        With tSelectedPics(bSelected).tActualSize
            .Left = X - pPos.X
            .Top = Y - pPos.Y
        End With
    Else
        With tSelectedPics(bSelected).tActualSize
            Select Case sEdge
                Case "LEFT"
                    .Right = .Right + .Left - X
                    .Left = X
                Case "RIGHT": .Right = X - .Left
                Case "TOP"
                    .Bottom = .Bottom + .Top - Y
                    .Top = Y
                Case "BOTTOM": .Bottom = Y - .Top
                Case "TOPLEFT":
                    .Right = .Right + .Left - X
                    .Left = X
                    .Bottom = .Bottom + .Top - Y
                    .Top = Y
                Case "TOPRIGHT"
                    .Bottom = .Bottom + .Top - Y
                    .Top = Y
                    .Right = X - .Left
                Case "BOTTOMLEFT"
                    .Right = .Right + .Left - X
                    .Left = X
                    .Bottom = Y - .Top
                Case "BOTTOMRIGHT"
                    .Bottom = Y - .Top
                    .Right = X - .Left
            End Select

            If .Bottom < 0 Then
                .Top = .Top + .Bottom
                .Bottom = Abs(.Bottom)

                If tSelectedPics(bSelected).iValueY = 1 Then
                    tSelectedPics(bSelected).iValueY = -1
                Else: tSelectedPics(bSelected).iValueY = 1
                End If

                If sEdge = "BOTTOM" Then
                    sEdge = "TOP"
                ElseIf sEdge = "TOP" Then
                    sEdge = "BOTTOM"
                ElseIf sEdge = "TOPLEFT" Then
                    sEdge = "BOTTOMLEFT"
                ElseIf sEdge = "BOTTOMLEFT" Then
                    sEdge = "TOPLEFT"
                ElseIf sEdge = "TOPRIGHT" Then
                    sEdge = "BOTTOMRIGHT"
                ElseIf sEdge = "BOTTOMRIGHT" Then
                    sEdge = "TOPRIGHT"
                End If
            End If

            If .Right < 0 Then
                .Left = .Left + .Right
                .Right = Abs(.Right)

                If tSelectedPics(bSelected).iValueX = 1 Then
                    tSelectedPics(bSelected).iValueX = -1
                Else: tSelectedPics(bSelected).iValueX = 1
                End If

                If sEdge = "RIGHT" Then
                    sEdge = "LEFT"
                ElseIf sEdge = "LEFT" Then
                    sEdge = "RIGHT"
                ElseIf sEdge = "TOPLEFT" Then
                    sEdge = "TOPRIGHT"
                ElseIf sEdge = "TOPRIGHT" Then
                    sEdge = "TOPLEFT"
                ElseIf sEdge = "BOTTOMRIGHT" Then
                    sEdge = "BOTTOMLEFT"
                ElseIf sEdge = "BOTTOMLEFT" Then
                    sEdge = "BOTTOMRIGHT"
                End If
            End If
        End With
    End If

    DrawPage
Else
    bTemp = bOver

    For nIndex = 1 To picShown.UBound
        With tSelectedPics(nIndex).tActualSize
            If X > .Left And X < .Left + .Right Then
                If Y > .Top And Y < .Top + .Bottom Then
                    bOverPic = nIndex
                    Exit For
                End If
            End If
        End With
    Next nIndex

    If bOverPic <> bTemp Then
        bOver = bOverPic
        DrawPage
    End If

    sEdge = CheckEdge(X, Y, bOver)
End If

If bSelected = 0 And sldOrder.Enabled = True Then sldOrder.Enabled = False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If bDown = 1 Then
    bDown = 0

    DrawPage
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub sldOrder_Change()

On Error GoTo ErrorHandler

Dim bNewOrder As Byte
Dim nIndex As Integer
Dim iLeft As Integer
Dim iTop As Integer
Dim iRight As Integer
Dim iBottom As Integer
Dim iStateX As Integer
Dim iStateY As Integer

If bByCode = True Then Exit Sub

bNewOrder = sldOrder.Max - (sldOrder.Value - 1)

If bNewOrder > bSelected Then
    For nIndex = bSelected To bNewOrder - 1
        With picShown(nIndex)
            picHidden.Width = .Width
            picHidden.Height = .Height
            BitBlt picHidden.hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
            .Width = picShown(nIndex + 1).Width
            .Height = picShown(nIndex + 1).Height
        End With

        With picShown(nIndex + 1)
            BitBlt picShown(nIndex).hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
            .Width = picHidden.Width
            .Height = picHidden.Height
        End With

        With picHidden
            BitBlt picShown(nIndex + 1).hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
        End With

        With tSelectedPics(nIndex).tActualSize
            iLeft = .Left
            iTop = .Top
            .Left = tSelectedPics(nIndex + 1).tActualSize.Left
            .Top = tSelectedPics(nIndex + 1).tActualSize.Top
            iRight = .Right
            iBottom = .Bottom
            .Right = tSelectedPics(nIndex + 1).tActualSize.Right
            .Bottom = tSelectedPics(nIndex + 1).tActualSize.Bottom
        End With

        iStateX = tSelectedPics(nIndex).iValueX
        iStateY = tSelectedPics(nIndex).iValueY
        tSelectedPics(nIndex).iValueX = tSelectedPics(nIndex + 1).iValueX
        tSelectedPics(nIndex).iValueY = tSelectedPics(nIndex + 1).iValueY
        tSelectedPics(nIndex + 1).iValueX = iStateX
        tSelectedPics(nIndex + 1).iValueY = iStateY

        With tSelectedPics(nIndex + 1).tActualSize
            .Left = iLeft
            .Top = iTop
            .Right = iRight
            .Bottom = iBottom
        End With
    Next nIndex

    bSelected = bNewOrder
Else
    For nIndex = bSelected To bNewOrder + 1 Step -1
        With picShown(nIndex)
            picHidden.Width = .Width
            picHidden.Height = .Height
            BitBlt picHidden.hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
            .Width = picShown(nIndex - 1).Width
            .Height = picShown(nIndex - 1).Height
        End With

        With picShown(nIndex - 1)
            BitBlt picShown(nIndex).hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
            .Width = picHidden.Width
            .Height = picHidden.Height
        End With

        With picHidden
            BitBlt picShown(nIndex - 1).hdc, 0, 0, .Width, .Height, .hdc, 0, 0, vbSrcCopy
        End With

        With tSelectedPics(nIndex).tActualSize
            iLeft = .Left
            iTop = .Top
            .Left = tSelectedPics(nIndex - 1).tActualSize.Left
            .Top = tSelectedPics(nIndex - 1).tActualSize.Top
            iRight = .Right
            iBottom = .Bottom
            .Right = tSelectedPics(nIndex - 1).tActualSize.Right
            .Bottom = tSelectedPics(nIndex - 1).tActualSize.Bottom
        End With

        iStateX = tSelectedPics(nIndex).iValueX
        iStateY = tSelectedPics(nIndex).iValueY
        tSelectedPics(nIndex).iValueX = tSelectedPics(nIndex - 1).iValueX
        tSelectedPics(nIndex).iValueY = tSelectedPics(nIndex - 1).iValueY
        tSelectedPics(nIndex - 1).iValueX = iStateX
        tSelectedPics(nIndex - 1).iValueY = iStateY

        With tSelectedPics(nIndex - 1).tActualSize
            .Left = iLeft
            .Top = iTop
            .Right = iRight
            .Bottom = iBottom
        End With
    Next nIndex

    bSelected = bNewOrder
End If

DrawPage

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function CheckEdge(ByVal X As Integer, ByVal Y As Integer, ByVal Index As Integer) As String

On Error GoTo ErrorHandler

Dim bNot As Byte

Const EDGE As Integer = 5

If Index = 0 Or Index <> bSelected Then picPage.MousePointer = 0: Exit Function

With tSelectedPics(Index).tActualSize
    X = X - .Left
    Y = Y - .Top

    If X >= 0 And X <= EDGE And Y >= EDGE And Y <= .Bottom - EDGE Then
        CheckEdge = "LEFT"

        If sEdge <> "LEFT" Then
            picPage.MouseIcon = LoadResPicture(25, vbResCursor)
        End If
    ElseIf X >= .Right - 5 And X <= .Right And Y >= EDGE And Y <= .Bottom - EDGE Then
        CheckEdge = "RIGHT"

        If sEdge <> "RIGHT" Then
            picPage.MouseIcon = LoadResPicture(25, vbResCursor)
        End If
    ElseIf X >= EDGE And X <= .Right - EDGE And Y >= 0 And Y <= EDGE Then
        CheckEdge = "TOP"

        If sEdge <> "TOP" Then
            picPage.MouseIcon = LoadResPicture(26, vbResCursor)
        End If
    ElseIf X >= EDGE And X <= .Right - EDGE And Y >= .Bottom - EDGE And Y <= .Bottom Then
        CheckEdge = "BOTTOM"

        If sEdge <> "BOTTOM" Then
            picPage.MouseIcon = LoadResPicture(26, vbResCursor)
        End If
    ElseIf X >= 0 And X <= EDGE And Y >= 0 And Y <= EDGE Then
        CheckEdge = "TOPLEFT"

        If sEdge <> "TOPLEFT" Then
            picPage.MouseIcon = LoadResPicture(27, vbResCursor)
        End If
    ElseIf X >= .Right - EDGE And X <= .Right And Y >= 0 And Y <= EDGE Then
        CheckEdge = "TOPRIGHT"

        If sEdge <> "TOPRIGHT" Then
            picPage.MouseIcon = LoadResPicture(28, vbResCursor)
        End If
    ElseIf X >= 0 And X <= EDGE And Y >= .Bottom - EDGE And Y <= .Bottom Then
        CheckEdge = "BOTTOMLEFT"

        If sEdge <> "BOTTOMLEFT" Then
            picPage.MouseIcon = LoadResPicture(28, vbResCursor)
        End If
    ElseIf X >= .Right - EDGE And X <= .Right And Y >= .Bottom - EDGE And Y <= .Bottom Then
        CheckEdge = "BOTTOMRIGHT"

        If sEdge <> "BOTTOMRIGHT" Then
            picPage.MouseIcon = LoadResPicture(27, vbResCursor)
        End If
    Else
        bNot = 1
        picPage.MousePointer = 0
    End If

    If bNot = 0 And picPage.MousePointer <> 99 Then picPage.MousePointer = 99
End With

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub sldOrder_Scroll()

On Error GoTo ErrorHandler

sldOrder_Change

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
