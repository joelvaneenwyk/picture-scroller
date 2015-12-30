VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Preview Pictures"
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   105
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lPictureIndex As Long
Private Sub Form_Click()

Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Unload Me

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim iWidth As Integer
Dim iHeight As Integer
Dim fScale As Single

Dim iMaxW As Integer
Dim iMaxH As Integer

Const LOADING = "Loading Picture..."

With Me
    .Width = (Me.TextWidth(LOADING) + 10) * Screen.TwipsPerPixelX
    .Height = (Me.TextHeight(LOADING) + 10) * Screen.TwipsPerPixelY

    .Cls

    Me.Line (0, 0)-(Me.ScaleWidth - 1, 0)
    Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1)
    Me.Line (Me.ScaleWidth - 1, Me.ScaleHeight - 1)-(0, Me.ScaleHeight - 1)
    Me.Line (0, Me.ScaleHeight - 1)-(0, 0)

    DrawTextAPI .hdc, LOADING, Len(LOADING), DirectDraw.GetRect(0, 0, .ScaleWidth, .ScaleHeight), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE

    .Show
    DoEvents
End With

GetPicFromIndex picTemp, lPictureIndex

iWidth = picTemp.Width
iHeight = picTemp.Height

iMaxW = (Screen.Width \ Screen.TwipsPerPixelX) - 20
iMaxH = (Screen.Height \ Screen.TwipsPerPixelY) - 20

If iWidth > iMaxW Or iHeight > iMaxH Then
    GetProportional iWidth, iHeight, iMaxW, iMaxH

    Me.Width = iMaxW * Screen.TwipsPerPixelX
    Me.Height = iMaxH * Screen.TwipsPerPixelY
Else
    Me.Width = iWidth * Screen.TwipsPerPixelX
    Me.Height = iHeight * Screen.TwipsPerPixelY
End If

iMaxW = Me.ScaleWidth
iMaxH = Me.ScaleHeight

GetProportional iWidth, iHeight, iMaxW, iMaxH

iWidth = iMaxW
iHeight = iMaxH

Me.Cls
SetStretchBltMode Me.hdc, STRETCH_DELETESCANS
StretchBlt Me.hdc, (Me.ScaleWidth - iWidth) / 2, (Me.ScaleHeight - iHeight) / 2, iWidth, iHeight, picTemp.hdc, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
Me.Refresh

Me.Show
DoEvents

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
