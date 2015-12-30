VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265358979

Dim iSpeed As Integer

Private Type RandomPoints
    X As Integer
    Y As Integer
End Type

' Note:
'   These transitions work by doing ONE frame (or step)
'   of the picture at one time.  So, the code will
'   draw one bit of the picture.  The next time it is
'   called, it will draw another bit, all the way until
'   it has printed the whole picture.  It can do this
'   because when it is first called, it sets "bOld" to
'   TRUE, meaning that we are now in the process of
'   drawing the transition.  When it's done, "bOld" is set
'   to FALSE: we're done.  Can be set from the outside.
Dim bOld As Byte

' Note:
'   "bAlwaysRedo" tells everyone that if the user
'   interrupts the transition half way (e.g., they
'   minimize Picture Scroller), then we need to start
'   from the beginning all over again.  Some transitions
'   need to do this, others do not.
Dim bAlwaysRedo As Byte
Public Sub MoveRD(ByVal bDirection As Byte)

Static iPositionX As Integer
Static iPositionY As Integer
Static iOriginalX As Integer
Static iOriginalY As Integer
Static iWidth As Integer
Static iHeight As Integer
Dim iTemp As Integer

Static Speed As Integer

If bOld = False Then
    bAlwaysRedo = False

    Speed = iScreenWidth / 500

    If bDirection = 0 Then
        iOriginalX = iPicWidth - Speed
        iPositionY = iPicTop
        iHeight = iPicHeight
    ElseIf bDirection = 1 Then
        iOriginalY = iPicHeight - Speed
        iPositionX = iPicLeft
        iWidth = iPicWidth
    End If

    bOld = True
End If

If bDirection = 0 And iPositionX > 0 Then iTemp = 1
If bDirection = 1 And iPositionY > 0 Then iTemp = 1

If iTemp = 1 Then
    If bDirection = 0 Then
        DirectDraw.ClearSurface dixuTransition, iPositionX - Speed, iPicTop, Speed, iPicHeight
    Else: DirectDraw.ClearSurface dixuTransition, iPicLeft, iPositionY - Speed, iPicWidth, Speed
    End If
    
    iTemp = 0
End If

If bDirection = 0 Then iWidth = iPicWidth - iOriginalX Else iHeight = iPicHeight - iOriginalY

DirectDraw.BltFast dixuTransition, iPositionX, iPositionY, iWidth, iHeight, dixuPicture, iOriginalX, iOriginalY, False

If bDirection = 0 And iOriginalX > 0 Then iTemp = 1
If bDirection = 1 And iOriginalY > 0 Then iTemp = 1

If iTemp = 1 Then
    If bDirection = 0 Then
        If iOriginalX - Speed < 0 Then
            iOriginalX = 0
        Else: iOriginalX = iOriginalX - Speed
        End If
    ElseIf bDirection = 1 Then
        If iOriginalY - Speed < 0 Then
            iOriginalY = 0
        Else: iOriginalY = iOriginalY - Speed
        End If
    End If
    
    iTemp = 0
Else
    If bDirection = 0 Then
        If iPositionX + Speed > iPicLeft Then
            iPositionX = iPicLeft
        Else: iPositionX = iPositionX + Speed
        End If
    ElseIf bDirection = 1 Then
        If iPositionY + Speed > iPicTop Then
            iPositionY = iPicTop
        Else: iPositionY = iPositionY + Speed
        End If
    End If
End If

If bDirection = 0 And iPositionX = iPicLeft And iOriginalX = 0 Then iTemp = 1
If bDirection = 1 And iPositionY = iPicTop And iOriginalY = 0 Then iTemp = 1

If iTemp = 1 Then
    iPositionX = 0: iOriginalX = 0
    iPositionY = 0: iOriginalY = 0

    frmScroller.TransitionDone
End If

End Sub
Public Sub MoveLU(ByVal bDirection As Byte)

Static iPositionX As Integer
Static iPositionY As Integer
Static iWidth As Integer
Static iHeight As Integer
Dim iTemp As Integer

Static Speed As Integer

If bOld = False Then
    bAlwaysRedo = False

    Speed = iScreenWidth / 500

    If Speed <= 0 Then Speed = 1

    If bDirection = 0 Then
        iWidth = Speed
        iHeight = iPicHeight
        iPositionY = iPicTop
        iPositionX = iScreenWidth - Speed
    ElseIf bDirection = 1 Then
        iHeight = Speed
        iWidth = iPicWidth
        iPositionX = iPicLeft
        iPositionY = iScreenHeight - Speed
    End If

    bOld = True
End If

If bDirection = 0 And iWidth > iPicWidth Then iWidth = iPicWidth
If bDirection = 1 And iHeight > iPicHeight Then iHeight = iPicHeight

If bDirection = 0 And iWidth >= iPicWidth Then iTemp = 1
If bDirection = 1 And iHeight >= iPicHeight Then iTemp = 1

If iTemp = 1 Then
    If bDirection = 0 Then
        If (iScreenWidth - iPositionX) - iPicWidth < Speed Then
            DirectDraw.ClearSurface dixuTransition, iPositionX + iPicWidth, iPicTop, (iScreenWidth - iPositionX) - iPicWidth, iPicHeight
        Else: DirectDraw.ClearSurface dixuTransition, iPositionX + iPicWidth, iPicTop, Speed, iPicHeight
        End If
    ElseIf bDirection = 1 Then
        If (iScreenHeight - iPositionY) - iPicHeight < Speed Then
            DirectDraw.ClearSurface dixuTransition, iPicLeft, iPositionY + iPicHeight, iPicWidth, (iScreenHeight - iPositionY) - iPicHeight
        Else: DirectDraw.ClearSurface dixuTransition, iPicLeft, iPositionY + iPicHeight, iPicWidth, Speed
        End If
    End If

    iTemp = 0
End If
    
DirectDraw.BltFast dixuTransition, iPositionX, iPositionY, iWidth, iHeight, dixuPicture, 0, 0, False

If bDirection = 0 And iPositionX > iPicLeft Then
    iPositionX = iPositionX - Speed

    If iWidth < iPicWidth Then
        iWidth = iWidth + Speed
    End If
End If

If bDirection = 1 And iPositionY > iPicTop Then
    iPositionY = iPositionY - Speed

    If iHeight < iPicHeight Then
        iHeight = iHeight + Speed
    End If
End If

If bDirection = 0 And iPositionX <= iPicLeft Then iTemp = 1
If bDirection = 1 And iPositionY <= iPicTop Then iTemp = 1

If iTemp = 1 Then
    iPositionX = 0
    iPositionY = 0
    iWidth = 0
    iHeight = 0

    frmScroller.TransitionDone
End If

End Sub
Public Sub Blinds(ByVal bType As Byte)

Static iStartPoints() As Integer
Static nIndex As Integer
Dim iTotal As Integer
Dim nPos As Integer

Const BLINDNUM = 10

If bType = 0 Then iTotal = iPicWidth \ BLINDNUM Else iTotal = iPicHeight \ BLINDNUM

If bOld = False Then
    bAlwaysRedo = True

    ReDim iStartPoints(0 To BLINDNUM)

    For nIndex = 0 To BLINDNUM
        If nIndex = 0 Then
            If bType = 0 Then iStartPoints(nIndex) = iPicLeft Else iStartPoints(nIndex) = iPicTop
        Else: iStartPoints(nIndex) = iStartPoints(nIndex - 1) + iTotal
        End If
    Next

    nIndex = 0

    bOld = True
End If

For nPos = 0 To BLINDNUM
    If iStartPoints(nPos) - iPicLeft <= iPicWidth - 1 And bType = 0 Then
        DirectDraw.BltFast dixuTransition, iStartPoints(nPos), iPicTop, 1, iPicHeight, dixuPicture, iStartPoints(nPos) - iPicLeft, 0, False
    ElseIf iStartPoints(nPos) - iPicTop <= iPicHeight - 1 And bType = 1 Then
        DirectDraw.BltFast dixuTransition, iPicLeft, iStartPoints(nPos), iPicWidth, 1, dixuPicture, 0, iStartPoints(nPos) - iPicTop, False
    End If

    iStartPoints(nPos) = iStartPoints(nPos) + 1
Next nPos

If nIndex = iTotal Then
    frmScroller.TransitionDone
    Exit Sub
Else: nIndex = nIndex + 1
End If

End Sub
Public Sub BoxOut()

Static iResult As Integer
Static iWidth As Integer
Static nIndex As Integer
Static iHeight As Integer
Static iPositionX As Integer
Static iPositionY As Integer

If bOld = False Then
    bAlwaysRedo = True

    nIndex = 0

    If iPicWidth > iPicHeight Then
        iHeight = 10
        iResult = (iPicHeight \ 2) - 5

        iWidth = ((iPicWidth \ 2) - iResult) * 2
    Else
        iWidth = 10
        iResult = (iPicWidth \ 2) - 5

        iHeight = ((iPicHeight \ 2) - iResult) * 2
    End If

    iResult = (iPicWidth - iWidth) \ 2
    iPositionX = (iPicWidth \ 2) - (iWidth \ 2) + iPicLeft
    iPositionY = (iPicHeight \ 2) - (iHeight \ 2) + iPicTop
    DirectDraw.BltFast dixuTransition, iPositionX, iPositionY, iWidth, iHeight, dixuPicture, iPositionX - iPicLeft, iPositionY - iPicTop, False

    bOld = True
End If

DirectDraw.BltFast dixuTransition, iPositionX - nIndex, iPositionY - nIndex, iWidth, 1, dixuPicture, iPositionX - iPicLeft - nIndex, iPositionY - iPicTop - nIndex, False
DirectDraw.BltFast dixuTransition, iPositionX - nIndex, iPositionY - nIndex + iHeight - 1, iWidth, 1, dixuPicture, iPositionX - iPicLeft - nIndex, iPositionY - nIndex + iHeight - iPicTop - 1, False
DirectDraw.BltFast dixuTransition, iPositionX - nIndex, iPositionY - nIndex, 1, iHeight, dixuPicture, iPositionX - iPicLeft - nIndex, iPositionY - iPicTop - nIndex, False
DirectDraw.BltFast dixuTransition, iPositionX - nIndex + iWidth, iPositionY - nIndex, 1, iHeight, dixuPicture, iPositionX - iPicLeft + iWidth - nIndex, iPositionY - iPicTop - nIndex, False

If iWidth < iPicWidth Then
    iWidth = iWidth + 2
End If

If iHeight < iPicHeight Then
    iHeight = iHeight + 2
End If

If nIndex = iResult - 2 Then
    nIndex = 0
    iWidth = 0
    iHeight = 0

    frmScroller.TransitionDone
Else: nIndex = nIndex + 1
End If

End Sub
Public Sub BoxIn()

Static iPosT As Integer
Static iPosL As Integer

If bOld = False Then
    bAlwaysRedo = True

    iPosT = 0
    iPosL = 0

    bOld = True
End If

If iPosT <= iPicHeight \ 2 Then
    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + iPosT, iPicWidth, 1, dixuPicture, 0, iPosT, False
End If

If iPosL <= iPicWidth \ 2 Then
    DirectDraw.BltFast dixuTransition, iPicLeft + iPosL, iPicTop, 1, iPicHeight, dixuPicture, iPosL, 0, False
End If

If iPosT < iPicHeight \ 2 Then
    iPosT = iPosT + 1
End If

If iPosL < iPicWidth \ 2 Then
    iPosL = iPosL + 1
End If

DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + iPicHeight - iPosT, iPicWidth, 1, dixuPicture, 0, iPicHeight - iPosT, False
DirectDraw.BltFast dixuTransition, iPicLeft + iPicWidth - iPosL, iPicTop, 1, iPicHeight, dixuPicture, iPicWidth - iPosL, 0, False

If iPosL >= iPicWidth \ 2 And iPosT >= iPicHeight \ 2 Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub SlideIn()

Static iPositionX As Integer
Dim iBottomSize As Integer
Dim nIndex As Integer
Dim iHalf As Integer

Const Speed As Integer = 10

If bOld = False Then
    bAlwaysRedo = True

    iPositionX = Speed

    bOld = True
End If

iHalf = iPicWidth \ 2

If iPositionX <= iPicWidth Then
    iBottomSize = iPicHeight - (iPicHeight \ 4) * 3
    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, iPositionX, iPicHeight \ 4, dixuPicture, iPicWidth - iPositionX, 0, False
    DirectDraw.BltFast dixuTransition, iPicLeft + iPicWidth - iPositionX, iPicTop + iPicHeight - iBottomSize, iPositionX, iBottomSize, dixuPicture, 0, iPicHeight - iBottomSize, False
End If

If iPositionX > iHalf Then
    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + (iPicHeight \ 4), iPositionX - iHalf, iPicHeight \ 4, dixuPicture, iPicWidth - (iPositionX - iHalf), iPicHeight \ 4, False
    DirectDraw.BltFast dixuTransition, iPicLeft + iPicWidth - (iPositionX - iHalf), iPicTop + ((iPicHeight \ 4) * 2), iPositionX - iHalf, iPicHeight \ 4, dixuPicture, 0, (iPicHeight \ 4) * 2, False
End If

If iPositionX = iPicWidth + (iPicWidth \ 2) Then
    iPositionX = 0

    frmScroller.TransitionDone
Else
    If iPositionX + Speed > iPicWidth + (iPicWidth \ 2) And iPositionX + Speed < iPicWidth + (iPicWidth \ 2) + Speed Then
        iPositionX = iPicWidth + (iPicWidth \ 2)
    ElseIf iPositionX + Speed > iPicWidth And iPositionX + Speed < iPicWidth + Speed Then
        iPositionX = iPicWidth
    Else: iPositionX = iPositionX + Speed
    End If
End If

End Sub
Public Sub Smear()

Static iStartPoints As Integer
Dim iSize As Integer

iSize = iPicWidth \ 400
iSize = iSize + 1

If bOld = False Then
    bAlwaysRedo = True

    iStartPoints = iSize

    bOld = True
End If

DirectDraw.Blt dixuTransition, iPicLeft, iPicTop, iPicWidth - iStartPoints, iPicHeight, dixuPicture, iPicWidth - iStartPoints, 0, iSize, iPicHeight, False
DirectDraw.BltFast dixuTransition, iPicWidth - iStartPoints + iPicLeft, iPicTop, iSize, iPicHeight, dixuPicture, iPicWidth - iStartPoints, 0, False

iStartPoints = iStartPoints + iSize

If iStartPoints >= iPicWidth - iSize Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub Slide(ByVal bType As Byte)

Static iPosX() As Integer
Static iEndPos As Integer
Static iStartPos As Integer
Static Increment As Integer
Dim nIndex As Integer
Dim iPosY As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If bOld = False Then
    bAlwaysRedo = True

    If iPicWidth > iPicHeight Then
        Increment = iPicWidth / 200
    Else: Increment = iPicHeight / 200
    End If

    If Increment = 0 Then Increment = 1

    ReDim iPosX(0 To iPicHeight / Increment)

    iStartPos = 0
    iEndPos = 0

    If bType = 1 Or bType = 3 Then
        For nIndex = 0 To UBound(iPosX)
            iPosX(nIndex) = iPicWidth - Increment
        Next nIndex
    End If

    bOld = True
End If

For nIndex = iStartPos To iEndPos
    If bType = 0 Or bType = 1 Then
        iPosY = nIndex * Increment
    Else: iPosY = iPicHeight - ((nIndex + 1) * Increment)
    End If

    iWidth = Increment
    iHeight = Increment

    If CheckSizes(iWidth, iHeight, iPosX(nIndex), iPosY) <> -1 Then
        DirectDraw.BltFast dixuTransition, iPosX(nIndex) + iPicLeft, iPosY + iPicTop, iWidth, iHeight, dixuPicture, iPosX(nIndex), iPosY, False
    End If

    If bType = 0 Or bType = 2 Then
        iPosX(nIndex) = iPosX(nIndex) + Increment
    Else: iPosX(nIndex) = iPosX(nIndex) - Increment
    End If

    If (bType = 0 Or bType = 2) And iPosX(nIndex) > iPicWidth Then iStartPos = iStartPos + 1
    If (bType = 1 Or bType = 3) And iPosX(nIndex) < -Increment Then iStartPos = iStartPos + 1
Next nIndex

If iEndPos < UBound(iPosX) Then iEndPos = iEndPos + 1

If iStartPos > iEndPos Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub SqueezeSlide(ByVal bType As Byte)

Static iPosX() As Integer
Static iEndPos As Integer
Static iStartPos As Integer
Static Increment As Integer
Dim nIndex As Integer
Dim iPosY As Integer
Dim iWidth As Integer
Dim iHeight As Integer
Dim iTempPosX As Integer

If bOld = False Then
    bAlwaysRedo = True

    If iPicWidth > iPicHeight Then
        Increment = iPicWidth / 150
    Else: Increment = iPicHeight / 150
    End If

    If Increment = 0 Then Increment = 1

    If bType = 0 Or bType = 1 Then
        ReDim iPosX(0 To (iPicHeight / Increment) / 2 + 2)
    Else: ReDim iPosX(0 To iPicHeight / Increment)
    End If

    iStartPos = 0
    iEndPos = 0

    If bType = 1 Or bType = 3 Then
        For nIndex = 0 To UBound(iPosX)
            iPosX(nIndex) = iPicWidth - Increment
        Next nIndex
    End If

    bOld = True
End If

For nIndex = iStartPos To iEndPos
    iPosY = (nIndex * Increment)

    iWidth = Increment
    iHeight = Increment

    If CheckSizes(iWidth, iHeight, iPosX(nIndex), iPosY) <> -1 Then
        DirectDraw.BltFast dixuTransition, iPosX(nIndex) + iPicLeft, iPosY + iPicTop, iWidth, iHeight, dixuPicture, iPosX(nIndex), iPosY, False
    End If
Next nIndex

For nIndex = iStartPos To iEndPos
    iPosY = iPicHeight - (nIndex * Increment)

    iWidth = Increment
    iHeight = Increment

    If bType = 0 Or bType = 1 Then
        iTempPosX = iPosX(nIndex)
    Else: iTempPosX = iPicWidth - iPosX(nIndex) - Increment
    End If

    If CheckSizes(iWidth, iHeight, iTempPosX, iPosY) <> -1 Then
        DirectDraw.BltFast dixuTransition, iTempPosX + iPicLeft, iPosY + iPicTop, iWidth, iHeight, dixuPicture, iTempPosX, iPosY, False
    End If

    If bType = 0 Or bType = 2 Then
        iPosX(nIndex) = iPosX(nIndex) + Increment
    Else: iPosX(nIndex) = iPosX(nIndex) - Increment
    End If

    If (bType = 0 Or bType = 2) And iPosX(nIndex) > iPicWidth Then iStartPos = iStartPos + 1
    If (bType = 1 Or bType = 3) And iPosX(nIndex) < -Increment Then iStartPos = iStartPos + 1
Next nIndex

If iEndPos < UBound(iPosX) Then iEndPos = iEndPos + 1

If (bType = 0 And iPosX(UBound(iPosX)) >= iPicWidth - Increment) Or (bType = 1 And iPosX(UBound(iPosX)) <= Increment) Then
    frmScroller.TransitionDone
ElseIf (bType = 2 And iPicWidth - iPosX(iEndPos) - Increment < iPosX(UBound(iPosX) - iEndPos)) Or (bType = 3 And iPicWidth - iPosX(iEndPos) - Increment > iPosX(UBound(iPosX) - iEndPos)) Then
    frmScroller.TransitionDone
End If

End Sub
Public Function CrossIn()

Static iPosX1 As Integer
Static iPosX2 As Integer
Dim iHeight As Integer
Dim iWidth As Integer
Dim iPosY1 As Integer
Dim iPosY2 As Integer
Dim iPosY As Integer

Const SIZE As Integer = 7

If bOld = False Then
    bAlwaysRedo = True

    iPosX2 = iPicWidth - 2

    bOld = True
End If

For iPosY = 0 To iPicHeight Step (iPicHeight \ SIZE) * 2
    iWidth = 2: iHeight = iPicHeight \ SIZE: iPosY1 = iPosY

    If CheckSizes(iWidth, iHeight, iPosX1, iPosY1) <> -1 Then
        DirectDraw.BltFast dixuTransition, iPicLeft + iPosX1, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX1, iPosY, False
    End If

    iWidth = 2: iHeight = iPicHeight \ SIZE: iPosY2 = iPosY + (iPicHeight \ SIZE)

    If CheckSizes(iWidth, iHeight, iPosX2, iPosY2) <> -1 Then
        DirectDraw.BltFast dixuTransition, iPicLeft + iPosX2, iPicTop + iPosY2, iWidth, iHeight, dixuPicture, iPosX2, iPosY2, False
    End If
Next iPosY

iPosX1 = iPosX1 + 2
iPosX2 = iPosX2 - 2

If iPosX2 < -2 Then
    iPosX1 = 0
    iPosX2 = 0
    frmScroller.TransitionDone
End If

End Function
Public Function RandomLines(ByVal bType As Byte)

Static Points As New Collection
Dim nIndex As Integer
Dim Random As Integer

Const Speed As Integer = 5

If bOld = False Then
    bAlwaysRedo = True

    If bType = 0 Then
        For nIndex = 0 To iPicWidth - 1
            Points.Add nIndex
        Next nIndex
    ElseIf bType = 1 Then
        For nIndex = 0 To iPicHeight - 1
            Points.Add nIndex
        Next nIndex
    End If

    bOld = True
End If

For nIndex = 1 To Speed
    If Points.Count = 0 Then Exit For

    Random = Rnd(1) * Points.Count

    If Random = 0 Then Random = 1

    If bType = 0 Then
        DirectDraw.BltFast dixuTransition, Points(Random) + iPicLeft, iPicTop, 1, iPicHeight, dixuPicture, Points(Random), 0, False
    Else: DirectDraw.BltFast dixuTransition, iPicLeft, Points(Random) + iPicTop, iPicWidth, 1, dixuPicture, 0, Points(Random), False
    End If

    Points.Remove Random
Next

If Points.Count = 0 Then
    frmScroller.TransitionDone
End If

End Function
Public Sub MoveOut(ByVal bType As Byte)

Static iSize As Integer
Dim iPos As Integer

If bOld = False Then
    bAlwaysRedo = True

    iSize = 1

    bOld = True
End If

If bType = 0 Then
    iPos = iPicWidth / 2

    If iPos - iSize >= 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPos - iSize, iPicTop, iSize, iPicHeight, dixuPicture, 0, 0, False

    Do Until iPos + iSize <= iPicWidth
        iPos = iPos - 1
    Loop

    DirectDraw.BltFast dixuTransition, iPicLeft + iPos, iPicTop, iSize, iPicHeight, dixuPicture, iPicWidth - iSize, 0, False
Else
    iPos = iPicHeight / 2

    If iPos - iSize >= 0 Then DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + iPos - iSize, iPicWidth, iSize, dixuPicture, 0, 0, False

    Do Until iPos + iSize <= iPicHeight
        iPos = iPos - 1
    Loop

    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + iPos, iPicWidth, iSize, dixuPicture, 0, iPicHeight - iSize, False
End If

iSize = iSize + 1

If (bType = 0 And iSize >= iPicWidth / 2) Or (bType = 1 And iSize >= iPicHeight / 2) Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub Snake(ByVal bType As Integer)

Static iPosX As Integer
Static iPosY As Integer
Static bDirection As Byte
Dim iWidth As Integer
Dim iHeight As Integer

Const SPEED1 As Integer = 20
Const SPEED2 As Integer = 50

If bOld = False Then
    bAlwaysRedo = True

    iPosX = 0
    iPosY = 0

    bOld = True
End If

If bType = 0 Then
    iWidth = SPEED1
    iHeight = SPEED2
ElseIf bType = 1 Then
    iWidth = SPEED2
    iHeight = SPEED1
End If

If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False

If bType = 0 Then
    If (bDirection = 0 And iPosX >= iPicWidth) Or (bDirection = 1 And iPosX <= 0) Then
        iPosY = iPosY + SPEED2
        If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False
    
        If bDirection = 0 Then
            bDirection = 1
        Else: bDirection = 0
        End If
    ElseIf bDirection = 0 Then
        iPosX = iPosX + SPEED1
    ElseIf bDirection = 1 Then
        iPosX = iPosX - SPEED1
    End If
ElseIf bType = 1 Then
    If (bDirection = 0 And iPosY >= iPicHeight) Or (bDirection = 1 And iPosY <= 0) Then
        iPosX = iPosX + SPEED2
        If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False
    
        If bDirection = 0 Then
            bDirection = 1
        Else: bDirection = 0
        End If
    ElseIf bDirection = 0 Then
        iPosY = iPosY + SPEED1
    ElseIf bDirection = 1 Then
        iPosY = iPosY - SPEED1
    End If
End If

If bType = 0 And (iPosY >= iPicHeight And iPosX >= iPicWidth And bDirection = 0) Or (iPosY >= iPicHeight - SPEED2 And iPosX <= -SPEED1 And bDirection = 1) Then
    frmScroller.TransitionDone
ElseIf bType = 1 And (iPosY >= iPicHeight And iPosX >= iPicWidth And bDirection = 0) Or (iPosY <= -SPEED2 And iPosX >= iPicWidth And bDirection = 1) Then
    frmScroller.TransitionDone
End If


End Sub
Public Sub Maze()

Static bDirection As Integer
Static iBottom As Integer
Static iRight As Integer
Static iPosX As Integer
Static iPosY As Integer
Static iLeft As Integer
Static iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If bOld = False Then
    bAlwaysRedo = True

    iPosX = 0
    iPosY = 0
    bDirection = 0
    iRight = iPicWidth - 51
    iBottom = iPicHeight - 51
    iTop = 50
    iLeft = 50

    bOld = True
End If

If bDirection = 0 Or bDirection = 2 Then
    iWidth = 10
    iHeight = 50
Else
    iWidth = 50
    iHeight = 10
End If

If iLeft - 50 >= iPicWidth / 2 Or iBottom + 50 <= iPicHeight / 2 Or iTop - 50 >= iPicWidth / 2 Or iRight + 50 <= iPicWidth / 2 Then GoTo EndSub

If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False

If bDirection = 0 Then
    If iPosX >= iRight Then
        bDirection = 1
        DirectDraw.BltFast dixuTransition, iPicLeft + iRight + 1, iPicTop + iPosY, 50, 50, dixuPicture, iRight + 1, iPosY, False
        iPosY = iPosY + 50
        iPosX = iRight + 1
        iRight = iRight - 50
    Else: iPosX = iPosX + 10
    End If
ElseIf bDirection = 1 Then
    If iPosY >= iBottom Then
        bDirection = 2
        DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iBottom + 1, 50, 50, dixuPicture, iPosX, iBottom + 1, False
        iPosX = iRight + 51
        iPosY = iBottom + 1
        iBottom = iBottom - 50
    Else: iPosY = iPosY + 10
    End If
ElseIf bDirection = 2 Then
    If iPosX <= iLeft Then
        bDirection = 3
        iPosX = iLeft - 50
        DirectDraw.BltFast dixuTransition, iPicLeft + iLeft - 50, iPicTop + iPosY, 50, 50, dixuPicture, iLeft - 50, iPosY, False
        iPosY = iBottom + 50
        iLeft = iLeft + 50
    Else: iPosX = iPosX - 10
    End If
ElseIf bDirection = 3 Then
    If iPosY <= iTop Then
        bDirection = 0
        iPosY = iTop
        DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iTop, 50, 50, dixuPicture, iPosX, iTop, False
        iPosX = iLeft - 50
        iTop = iTop + 50
    Else: iPosY = iPosY - 10
    End If
End If

Exit Sub

EndSub:
    frmScroller.TransitionDone

End Sub
Public Sub FallingLines(ByVal bType As Byte)

Static cPos1() As Integer
Static cPos2 As New Collection
Dim iRandom As Integer
Dim nIndex As Integer
Dim iWidth As Integer
Dim iHeight As Integer
Dim iTemp As Integer

Const Speed As Integer = 4

If bOld = False Then
    bAlwaysRedo = True

    If bType = 0 Or bType = 2 Then
        ReDim cPos1(0 To iPicWidth / Speed)

        For nIndex = 0 To (iPicWidth / Speed)
            If bType = 2 Then cPos1(nIndex) = iPicHeight - Speed
            cPos2.Add nIndex * Speed
        Next nIndex
    Else
        ReDim cPos1(0 To iPicHeight / Speed)

        For nIndex = 0 To (iPicHeight / Speed)
            If bType = 3 Then cPos1(nIndex) = iPicWidth - Speed
            cPos2.Add nIndex * Speed
        Next nIndex
    End If

    bOld = True
End If

For nIndex = 0 To 100
    If cPos2.Count = 0 Then Exit For

    If nIndex = 0 Then
        iRandom = cPos2.Count
    Else: iRandom = Rnd(1) * cPos2.Count
    End If

    If iRandom = 0 Then iRandom = 1

    iWidth = 4
    iHeight = 4

    iTemp = cPos2(iRandom) / Speed

    If bType = 0 Or bType = 2 Then
        If CheckSizes(iWidth, iHeight, cPos2(iRandom), cPos1(iTemp)) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + cPos2(iRandom), iPicTop + cPos1(iTemp), iWidth, iHeight, dixuPicture, cPos2(iRandom), cPos1(iTemp), False
    Else: If CheckSizes(iWidth, iHeight, cPos1(iTemp), cPos2(iRandom)) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + cPos1(iTemp), iPicTop + cPos2(iRandom), iWidth, iHeight, dixuPicture, cPos1(iTemp), cPos2(iRandom), False
    End If

    If (bType = 0 And cPos1(iTemp) >= iPicHeight - Speed) Or (bType = 1 And cPos1(iTemp) >= iPicWidth - Speed) Or (bType = 2 And cPos1(iTemp) <= 0) Or (bType = 3 And cPos1(iTemp) <= 0) Then
        cPos2.Remove iRandom
    Else
        If bType = 0 Or bType = 1 Then
            cPos1(iTemp) = cPos1(iTemp) + Speed
        Else: cPos1(iTemp) = cPos1(iTemp) - Speed
        End If
    End If
Next nIndex

If cPos2.Count = 0 Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub Wipe(ByVal bType As Byte)

Static iPosX As Integer
Static iPosY As Integer
Static iWidth As Integer
Static iHeight As Integer

Const Speed As Integer = 5

If bOld = False Then
    bAlwaysRedo = True

    Select Case bType
        Case 0: iWidth = Speed: iPosX = 0: iHeight = iPicHeight: iPosY = 0
        Case 1: iWidth = iPicWidth: iPosX = 0: iHeight = Speed: iPosY = 0
        Case 2: iWidth = Speed: iPosX = iPicWidth - Speed: iHeight = iPicHeight: iPosY = 0
        Case 3: iWidth = iPicWidth: iPosX = 0: iHeight = Speed: iPosY = iPicHeight - Speed
    End Select

    bOld = True
End If

If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False

Select Case bType
    Case 0: iPosX = iPosX + Speed
    Case 1: iPosY = iPosY + Speed
    Case 2: iPosX = iPosX - Speed
    Case 3: iPosY = iPosY - Speed
End Select

If (bType = 0 And iPosX >= iPicWidth) Or (bType = 1 And iPosY >= iPicHeight) Or (bType = 2 And iPosX < 0) Or (bType = 3 And iPosY < 0) Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub StretchOut(ByVal bType As Byte)

Static iWidth As Integer
Static iHeight As Integer
Dim iPosX As Integer
Dim iPosY As Integer

If bOld = False Then
    bAlwaysRedo = True

    iWidth = iPicWidth / 50
    iHeight = iPicHeight / 50

    bOld = True
End If

iPosX = (iPicWidth - iWidth) / 2
iPosY = (iPicHeight - iHeight) / 2

If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False

If bType = 0 Then
    If iPosX > 0 Then
        iWidth = iWidth + 15
    ElseIf iPosY > 0 Then
        iHeight = iHeight + 15
    End If
ElseIf bType = 1 Then
    If iPosY > 0 Then
        iHeight = iHeight + 15
    ElseIf iPosX > 0 Then
        iWidth = iWidth + 15
    End If
End If

If (bType = 0 And iPosY = 0) Or (bType = 1 And iPosX = 0) Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub Shades(ByVal bType As Byte)

Static iPosY As Integer
Static iPosX As Integer
Static Speed As Integer
Static iSize As Integer
Static bStart As Byte
Static bTimes As Byte
Dim nIndex As Integer
Dim iWidth As Integer
Dim iHeight As Integer
Dim iTimes As Integer

If bOld = False Then
    bAlwaysRedo = True

    If bType = 0 Or bType = 1 Then
        iSize = iPicWidth
    Else: iSize = iPicHeight
    End If

    Speed = iSize / 200

    If Speed = 0 Then Speed = 1

    iPosX = 0
    iPosY = 0
    bStart = 0
    bTimes = 0

    bOld = True
End If

For iTimes = 0 To 1
    For nIndex = bStart To iSize - Speed Step Speed * 2
        iWidth = Speed
        iHeight = Speed

        If bType = 0 Or bType = 1 Then
            iPosX = nIndex
        Else: iPosY = nIndex
        End If

        If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False
    Next nIndex

    If bStart = 0 Then
        bStart = Speed
    Else: bStart = 0
    End If

    If (bType = 0 Or bType = 1) And ((bTimes = 0 Or (bTimes = 1 And bType = 0)) And iPosY > iPicHeight - Speed) Or (bType = 1 And iPosY < 0) Then
        bTimes = bTimes + 1

        If bType = 0 Then
            bStart = Speed
            iPosY = 0
        End If
    ElseIf (bType = 2 Or bType = 3) And ((bTimes = 0 Or (bTimes = 1 And bType = 2)) And iPosX > iPicWidth - Speed) Or (bType = 3 And iPosX < 0) Then
        bTimes = bTimes + 1

        If bType = 2 Then
            bStart = Speed
            iPosX = 0
        End If
    ElseIf bTimes = 0 Or bType = 0 Or bType = 2 Then
        If bType = 0 Or bType = 1 Then
            iPosY = iPosY + Speed
        ElseIf bType = 2 Or bType = 3 Then
            iPosX = iPosX + Speed
        End If
    ElseIf bTimes = 1 Then
        If bType = 1 Then
            iPosY = iPosY - Speed
        ElseIf bType = 3 Then
            iPosX = iPosX - Speed
        End If
    End If

    If bTimes = 2 Then
        frmScroller.TransitionDone
    End If
Next iTimes

End Sub
Public Sub Puzzle()

Static iHeight As Integer
Static iTHeight As Integer
Dim iWidth As Integer
Dim iPosX As Integer
Dim iPosY As Integer
Dim SIZE As Integer

SIZE = iPicHeight - (iPicHeight / 2)

Const Speed As Integer = 2

If bOld = False Then
    bAlwaysRedo = True

    iHeight = Speed

    bOld = True
End If

If iPicHeight - iHeight < SIZE Then iHeight = iPicHeight - SIZE

iWidth = iPicWidth / 2: iPosY = SIZE - iHeight: iTHeight = iHeight
If CheckSizes(iWidth, iTHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iTHeight, dixuPicture, iPosX, 0, False
iWidth = iPicWidth / 2: iPosY = SIZE: iTHeight = iHeight
If CheckSizes(iWidth, iTHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPicHeight - iHeight, iWidth, iTHeight, dixuPicture, iPosX, iPosY, False
iWidth = iPicWidth - (iPicWidth / 2): iPosY = SIZE - iHeight: iTHeight = iHeight: iPosX = iPicWidth / 2
If CheckSizes(iWidth, iTHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop, iWidth, iTHeight, dixuPicture, iPosX, iPosY, False
iWidth = iPicWidth - (iPicWidth / 2): iPosY = SIZE: iTHeight = iHeight
If CheckSizes(iWidth, iTHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + SIZE, iWidth, iTHeight, dixuPicture, iPosX, iPicHeight - iHeight, False

If iPicHeight - iHeight <= SIZE Then
    frmScroller.TransitionDone
End If

iHeight = iHeight + Speed

End Sub
Public Sub Sandwich(ByVal bType As Byte)

Static iSize As Integer

Const Speed As Integer = 4

If bOld = False Then
    bAlwaysRedo = True

    iSize = Speed

    bOld = True
End If

If bType = 0 Then
    DirectDraw.BltFast dixuTransition, iPicLeft + iSize - Speed, iPicTop, Speed, iPicHeight, dixuPicture, iSize - Speed, 0, False
    DirectDraw.BltFast dixuTransition, iPicLeft + iPicWidth - iSize, iPicTop, Speed, iPicHeight, dixuPicture, iPicWidth - iSize, 0, False
Else
    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + iSize - Speed, iPicWidth, Speed, dixuPicture, 0, iSize - Speed, False
    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop + iPicHeight - iSize, iPicWidth, Speed, dixuPicture, 0, iPicHeight - iSize, False
End If

If (bType = 0 And iSize - Speed > iPicWidth - iSize - Speed) Or (bType = 1 And iSize - Speed > iPicHeight - iSize - Speed) Then
    frmScroller.TransitionDone
End If

iSize = iSize + Speed

End Sub
Public Sub TileAcross()

Static iPosX As Integer
Static iOriginalX As Integer
Dim iWidth As Integer
Dim iHeight As Integer
Dim iPosY As Integer
Dim SIZE As Integer
Dim Speed As Integer

SIZE = iPicWidth / 25
Speed = iPicWidth / 45

If bOld = False Then
    bAlwaysRedo = True

    iPosX = 0

    iOriginalX = iPicWidth - SIZE

    bOld = True
End If

iWidth = SIZE
iHeight = iPicHeight

If iPosX >= Speed Then DirectDraw.ClearSurface dixuTransition, iPicLeft + iPosX - Speed, iPicTop, Speed, iPicHeight

If CheckSizes(iWidth, iHeight, iOriginalX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iOriginalX, iPosY, False

If iOriginalX = 0 Then
    frmScroller.TransitionDone
    Exit Sub
End If

If iPosX = iOriginalX Then
    iOriginalX = iOriginalX - SIZE

    iPosX = 0

    If iOriginalX < 0 Then
        iOriginalX = 0
    End If
Else
    iPosX = iPosX + Speed

    If iPosX > iOriginalX Then
        iPosX = iOriginalX
    End If

    If iPosX < 0 Then iPosX = 0
End If

End Sub
Public Sub CircleOut()

Static iRadius As Integer
Dim iX(0 To 360) As Integer
Dim iY(0 To 360) As Integer
Dim Radians As Double
Dim nIndex As Integer
Dim iPosX As Integer
Dim iPosY As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If bOld = False Then
    bAlwaysRedo = True

    iRadius = 0

    bOld = True
End If

For Radians = 0 To 3.12413936106985 Step 0.017453292519
    iX(nIndex) = (iRadius * Cos(Radians)) + (0 * Sin(Radians))
    iY(nIndex) = (0 * Cos(Radians)) - (iRadius * Sin(Radians))
    iX(nIndex + 1) = (iRadius * Cos(6.28318530717959 - Radians)) + (0 * Sin(6.28318530717959 - Radians))
    iY(nIndex + 1) = (0 * Cos(6.28318530717959 - Radians)) - (iRadius * Sin(6.28318530717959 - Radians))
    nIndex = nIndex + 2
    DoEvents
Next

For nIndex = 0 To UBound(iX) - 1 Step 2
    iPosX = iX(nIndex) + (iPicWidth / 2)
    If iPosX < 0 Then iPosX = 0
    If iPosX > iPicWidth - 1 Then iPosX = iPicWidth - 1
    iPosY = iY(nIndex) + (iPicHeight / 2)
    If iPosY < 0 Then iPosY = 0
    If iPosY > iPicHeight - 1 Then iPosY = iPicHeight - 1
    If nIndex <> 0 Then iWidth = (iX(nIndex - 2) + (iPicWidth / 2)) - iPosX
    iHeight = (iY(nIndex + 1) + (iPicHeight / 2)) - iPosY
    If iWidth = 0 Then iWidth = 1
    If iHeight = 0 Then iHeight = 1
    If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False
Next nIndex

iRadius = iRadius + 5

If iRadius >= Sqr(((iPicWidth / 2) * (iPicWidth / 2)) + ((iPicHeight / 2) * (iPicHeight / 2))) Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub CircleIn()

Static iRadius As Integer
Dim iX(0 To 360) As Integer
Dim iY(0 To 360) As Integer
Dim Radians As Double
Dim nIndex As Integer
Dim iPosX As Integer
Dim iPosY As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If bOld = False Then
    bAlwaysRedo = True

    iRadius = Sqr(((iPicWidth / 2) * (iPicWidth / 2)) + ((iPicHeight / 2) * (iPicHeight / 2)))

    bOld = True
End If

DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, iPicWidth, iPicHeight, dixuPicture, 0, 0, False

For Radians = 0 To 3.12413936106985 Step 0.017453292519
    iX(nIndex) = (iRadius * Cos(Radians)) + (0 * Sin(Radians))
    iY(nIndex) = (0 * Cos(Radians)) - (iRadius * Sin(Radians))
    iX(nIndex + 1) = (iRadius * Cos(6.28318530717959 - Radians)) + (0 * Sin(6.28318530717959 - Radians))
    iY(nIndex + 1) = (0 * Cos(6.28318530717959 - Radians)) - (iRadius * Sin(6.28318530717959 - Radians))
    nIndex = nIndex + 2
    DoEvents
Next

For nIndex = 0 To UBound(iX) - 1 Step 2
    iPosX = iX(nIndex) + (iPicWidth / 2)
    If iPosX < 0 Then iPosX = 0
    If iPosX > iPicWidth - 1 Then iPosX = iPicWidth - 1
    iPosY = iY(nIndex) + (iPicHeight / 2)
    If iPosY < 0 Then iPosY = 0
    If iPosY > iPicHeight - 1 Then iPosY = iPicHeight - 1
    If nIndex <> 0 Then iWidth = (iX(nIndex - 2) + (iPicWidth / 2)) - iPosX
    iHeight = (iY(nIndex + 1) + (iPicHeight / 2)) - iPosY
    If iWidth = 0 Then iWidth = 1
    If iHeight = 0 Then iHeight = 1
    If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.ClearSurface dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight
Next nIndex

iRadius = iRadius - 5

If iRadius <= 1 Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub ArrangePieces()

Static cPosX As New Collection
Static cPosY As New Collection
Static cHit As New Collection
Dim nIndex1 As Integer
Dim nIndex2 As Integer
Static Speed As Integer
Dim iRandom As Integer
Dim iOrgRandom As Integer
Dim iWidth As Integer
Dim iHeight As Integer
Dim iTemp As Integer

If bOld = False Then
    bAlwaysRedo = True

    If iScreenWidth > iScreenHeight Then
        Speed = iScreenWidth \ 15
    Else: Speed = iScreenHeight \ 15
    End If

    If Speed <= 0 Then Speed = 1

    For nIndex1 = 0 To iPicHeight Step Speed
        For nIndex2 = 0 To iPicWidth + Speed Step Speed
            cPosX.Add nIndex2
            cPosY.Add nIndex1
            cHit.Add 0
        Next nIndex2
    Next nIndex1

    bOld = True
End If

iRandom = Rnd(1) * cHit.Count
If iRandom = 0 Then iRandom = 1
iOrgRandom = Rnd(1) * cHit.Count
If iOrgRandom = 0 Then iOrgRandom = 1

iWidth = Speed
iHeight = Speed

If cHit(iRandom) = 2 Then iOrgRandom = iRandom

If CheckSizes(iWidth, iHeight, cPosX(iRandom), cPosY(iRandom)) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + cPosX(iRandom), iPicTop + cPosY(iRandom), iWidth, iHeight, dixuPicture, cPosX(iOrgRandom), cPosY(iOrgRandom), False

If cHit(iRandom) = 2 Then
    cPosX.Remove iRandom
    cPosY.Remove iRandom
    cHit.Remove iRandom
Else
    iTemp = cHit(iRandom) + 1
    cHit.Remove iRandom

    If cHit.Count = 0 Then
        DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, iPicWidth, iPicHeight, dixuPicture, 0, 0, False
    ElseIf iRandom - 1 = cHit.Count Then
        cHit.Add iTemp, , , iRandom - 1
    Else: cHit.Add iTemp, , iRandom
    End If
End If

If cHit.Count = 0 Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub Cubes(ByVal bType As Byte)

Static lPos() As Long
Static nIndex As Long
Dim lTemp As Long
Dim lPoint1 As Long
Dim lPoint2 As Long
Dim nIn As Integer

Static SIZE As Single

If bOld = False Then
    bAlwaysRedo = True

    If iPicWidth > iPicHeight Then
        SIZE = iPicWidth / (500 / ((bType + 1) ^ 2))
    Else: SIZE = iPicHeight / (500 / ((bType + 1) ^ 2))
    End If

    If SIZE > 1 And SIZE < 2 Then SIZE = 2

    If SIZE <= 0 Then SIZE = 1

    SIZE = Round(SIZE)

    nIndex = 0

    ReDim lPos(0 To (iPicWidth / SIZE) * (iPicHeight / SIZE))

    DirectDraw.BltFast dixuTransition, iPicLeft, iPicTop, SIZE, SIZE, dixuPicture, 0, 0, False

    bOld = True
End If

For nIndex = nIndex To nIndex + 500
    If nIndex Mod 500 = 0 Or nIndex = 0 Then
        For nIn = 0 To 500
            If nIndex + nIn = UBound(lPos) Then Exit For

            If lPos(nIndex + nIn) = 0 Then
                lPos(nIndex + nIn) = nIndex + nIn
            End If

            lPoint1 = Rnd(1) * (UBound(lPos) - nIndex) + nIndex

            If lPos(lPoint1) = 0 Then lPos(lPoint1) = lPoint1

            lTemp = lPos(nIndex + nIn)

            lPos(nIndex + nIn) = lPos(lPoint1)
            lPos(lPoint1) = lTemp
        Next nIn
    End If

    lPoint2 = Int(lPos(nIndex) / Int((iPicWidth / SIZE)))

    lPoint1 = lPos(nIndex) - lPoint2 * Int((iPicWidth / SIZE))

    lPoint1 = lPoint1 * SIZE
    lPoint2 = lPoint2 * SIZE

    DirectDraw.BltFast dixuTransition, iPicLeft + lPoint1, iPicTop + lPoint2, SIZE, SIZE, dixuPicture, lPoint1, lPoint2, False

    If nIndex = UBound(lPos) Then Exit For
Next nIndex

If nIndex = UBound(lPos) Then
    frmScroller.TransitionDone
End If

End Sub
Private Function CheckSizes(iWidth As Integer, iHeight As Integer, iPosX As Integer, iPosY As Integer) As Integer

If iPosX = iPicWidth Then CheckSizes = -1

If iPosX < 0 And iPosX + iWidth > 0 Then
    iWidth = iWidth + iPosX
    iPosX = 0
End If

If iPosX + iWidth > iPicWidth And iPicWidth - iPosX > 0 Then
    iWidth = iPicWidth - iPosX
End If

If iPicWidth - iPosX < 0 Or iWidth + iPosX <= 0 Then CheckSizes = -1

If iPosY = iPicHeight Then CheckSizes = -1

If iPosY < 0 And iPosY + iHeight > 0 Then
    iHeight = iHeight + iPosY
    iPosY = 0
End If

If iPosY + iHeight > iPicHeight And iPicHeight - iPosY > 0 Then
    iHeight = iPicHeight - iPosY
End If

If iPicHeight - iPosY < 0 Or iHeight + iPosY <= 0 Then CheckSizes = -1

End Function
Public Property Get AlwaysRedo() As Byte

AlwaysRedo = bAlwaysRedo

End Property
Public Sub StopTransition()

bOld = False

End Sub
