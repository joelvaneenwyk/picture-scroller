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
Public Sub MoveLU()

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

    iWidth = Speed
    iHeight = iPicHeight
    iPositionY = iPicTop
    iPositionX = iScreenWidth - Speed

    bOld = True
End If

If iWidth > iPicWidth Then iWidth = iPicWidth

If iWidth >= iPicWidth Then iTemp = 1

If iTemp = 1 Then
    If (iScreenWidth - iPositionX) - iPicWidth < Speed Then
        DirectDraw.ClearSurface dixuTransition, iPositionX + iPicWidth, iPicTop, (iScreenWidth - iPositionX) - iPicWidth, iPicHeight
    Else: DirectDraw.ClearSurface dixuTransition, iPositionX + iPicWidth, iPicTop, Speed, iPicHeight
    End If

    iTemp = 0
End If
    
DirectDraw.BltFast dixuTransition, iPositionX, iPositionY, iWidth, iHeight, dixuPicture, 0, 0, False

If iPositionX > iPicLeft Then
    iPositionX = iPositionX - Speed

    If iWidth < iPicWidth Then
        iWidth = iWidth + Speed
    End If
End If

If iPositionX <= iPicLeft Then iTemp = 1

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

Static Speed As Integer
Static Increment As Integer

If bOld = False Then
    bAlwaysRedo = True

    If iPicWidth > iPicHeight Then
        Speed = iPicWidth / 20
        Increment = iPicWidth / 100
    Else
        Speed = iPicHeight / 20
        Increment = iPicHeight / 100
    End If

    If Speed = 0 Then Speed = 1
    If Increment = 0 Then Increment = 1

    iPosX = 0
    iPosY = 0
    bDirection = 0
    iRight = iPicWidth - Speed - 1
    iBottom = iPicHeight - Speed - 1
    iTop = Speed
    iLeft = Speed

    bOld = True
End If

If bDirection = 0 Or bDirection = 2 Then
    iWidth = Increment
    iHeight = Speed
Else
    iWidth = Speed
    iHeight = Increment
End If

If iLeft - Speed >= iPicWidth / 2 Or iBottom + Speed <= iPicHeight / 2 Or iTop - Speed >= iPicWidth / 2 Or iRight + Speed <= iPicWidth / 2 Then GoTo EndSub

If CheckSizes(iWidth, iHeight, iPosX, iPosY) = 0 Then DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iPosY, iWidth, iHeight, dixuPicture, iPosX, iPosY, False

If bDirection = 0 Then
    If iPosX >= iRight Then
        bDirection = 1
        DirectDraw.BltFast dixuTransition, iPicLeft + iRight + 1, iPicTop + iPosY, Speed, Speed, dixuPicture, iRight + 1, iPosY, False
        iPosY = iPosY + Speed
        iPosX = iRight + 1
        iRight = iRight - Speed
    Else: iPosX = iPosX + Increment
    End If
ElseIf bDirection = 1 Then
    If iPosY >= iBottom Then
        bDirection = 2
        DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iBottom + 1, Speed, Speed, dixuPicture, iPosX, iBottom + 1, False
        iPosX = iRight + Speed + 1
        iPosY = iBottom + 1
        iBottom = iBottom - Speed
    Else: iPosY = iPosY + Increment
    End If
ElseIf bDirection = 2 Then
    If iPosX <= iLeft Then
        bDirection = 3
        iPosX = iLeft - Speed
        DirectDraw.BltFast dixuTransition, iPicLeft + iLeft - Speed, iPicTop + iPosY, Speed, Speed, dixuPicture, iLeft - Speed, iPosY, False
        iPosY = iBottom + Speed
        iLeft = iLeft + Speed
    Else: iPosX = iPosX - Increment
    End If
ElseIf bDirection = 3 Then
    If iPosY <= iTop Then
        bDirection = 0
        iPosY = iTop
        DirectDraw.BltFast dixuTransition, iPicLeft + iPosX, iPicTop + iTop, Speed, Speed, dixuPicture, iPosX, iTop, False
        iPosX = iLeft - Speed
        iTop = iTop + Speed
    Else: iPosY = iPosY - Increment
    End If
End If

Exit Sub

EndSub:
    frmScroller.TransitionDone

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

Static Increment As Integer

If bOld = False Then
    bAlwaysRedo = True

    If iPicWidth > iPicHeight Then
        Increment = iPicWidth / 300
    Else: Increment = iPicHeight / 300
    End If

    If Increment = 0 Then Increment = 1

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

iRadius = iRadius - Increment

If iRadius <= 1 Then
    frmScroller.TransitionDone
End If

End Sub
Public Sub Cubes()

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
        SIZE = iPicWidth / 500
    Else: SIZE = iPicHeight / 500
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
