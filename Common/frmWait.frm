VERSION 5.00
Begin VB.Form frmWait 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Wait"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   197
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   600
   End
   Begin VB.PictureBox picAnimation 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2520
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer tmrDots 
      Interval        =   500
      Left            =   120
      Top             =   105
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   390
      Left            =   810
      TabIndex        =   0
      Top             =   643
      Width           =   1335
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Const COLOR_RANGE = &HDEDEDE

Dim POSX As Integer
Dim POSY As Integer

Public sWords As String
Public bCancelEnabled As Byte

Dim rRect As RECT
Dim bDotStep As Byte
Dim bPicStep As Byte
Dim bDirection As Byte
Private Sub cmdCancel_Click()

' The user wants to cancel!!!
bCancelOp = True

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim sString As String

sString = sWords & "..."

With rRect
    .Top = (cmdCancel.Top - Me.TextHeight(sString)) / 2
    .Bottom = .Top + Me.TextHeight(sString)
    .Left = (Me.ScaleWidth - Me.TextWidth(sString)) / 2
    .Right = .Left + Me.TextWidth(sString)
End With

tmrDots_Timer

POSX = (rRect.Left - 20) / 2
POSY = (Me.TextHeight(sString) - 20) / 2 + rRect.Top

tmrAnimation_Timer

cmdCancel.Enabled = bCancelEnabled

tmrDots.Enabled = True
tmrAnimation.Enabled = True

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrorHandler

tmrDots.Enabled = False

bDotStep = 0
bPicStep = 0
bDirection = 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrAnimation_Timer()

On Error GoTo ErrorHandler

Dim iPosX As Integer
Dim iPosY As Integer
Dim lColor As Long

Line (POSX, POSY)-(POSX + 20, POSY + 20), Me.BackColor, BF

If bPicStep = 5 Then
    bPicStep = 1
Else: bPicStep = bPicStep + 1
End If

picAnimation.Picture = LoadResPicture(bPicStep, vbResBitmap)

For iPosX = POSX To POSX + 19
    For iPosY = POSY To POSY + 19
        lColor = GetPixel(picAnimation.hdc, iPosX - POSX, iPosY - POSY)
        If lColor < COLOR_RANGE Then SetPixel Me.hdc, iPosX, iPosY, lColor
    Next iPosY
Next iPosX

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrDots_Timer()

On Error GoTo ErrorHandler

Dim sString As String

If bDotStep = 3 Then bDotStep = 0
bDotStep = bDotStep + 1

Line (rRect.Left, rRect.Top)-(rRect.Right, rRect.Bottom), Me.BackColor, BF

sString = sWords & String(bDotStep, ".")

DrawTextAPI Me.hdc, sString, Len(sString), rRect, DT_LEFT

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
