VERSION 5.00
Begin VB.Form frmWallpaper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Background"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   Icon            =   "frmWallpaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2198
      TabIndex        =   6
      Top             =   1980
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   668
      TabIndex        =   5
      Top             =   1980
      Width           =   1185
   End
   Begin VB.OptionButton optLocation 
      Caption         =   "Use Picture From Current Location"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   1575
      Width           =   2955
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1155
      Width           =   1935
   End
   Begin VB.OptionButton optLocation 
      Caption         =   "Copy Picture to Windows Directory"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   2955
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   3810
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   128
      X2              =   3908
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblTip 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Move your mouse over something to see what it does."
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   3810
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInstructions 
      AutoSize        =   -1  'True
      Caption         =   "Choose whether you want to copy this picture into your Windows directory or use it from its current location."
      Height          =   585
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3840
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "New File Name:"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1200
      Width           =   1125
   End
End
Attribute VB_Name = "frmWallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bDone As Byte

Dim sPictureFile As String

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Sub cmdCancel_Click()

bDone = True

End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If lblInfo.Tag <> "4" Then
    lblInfo.Caption = "Clicking this with cancel making this picture your background."
    lblInfo.Tag = 4
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdOK_Click()

On Error GoTo ErrorHandler

Dim sDirectory As String
Dim lReturnVal As Long
Dim sFileName As String

Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2

If optLocation(0).Value = True Then
    If txtFileName.Text = "" Then
        txtFileName.SetFocus

        lblInfo.Caption = "Please enter a file name to use for the copy of this picture that will be put in the Windows directory."
        tmrFlash.Enabled = True
        tmrFlash_Timer
        Exit Sub
    End If

    sDirectory = Space(255)
    lReturnVal = GetWindowsDirectory(sDirectory, 255)
    sDirectory = Left$(sDirectory, lReturnVal)
    NormalizePath sDirectory
    sFileName = sDirectory & txtFileName.Text & ".bmp"

    If FileExists(sFileName) = True Then
        txtFileName.SetFocus

        lblInfo.Caption = "A file with this name already exists.  Please enter a different file name."
        tmrFlash.Enabled = True
        tmrFlash_Timer
        Exit Sub
    End If

    GetPicFromIndex picTemp, frmScroller.lPictureIndex

    SavePicture picTemp.Picture, sFileName
Else: sFileName = sPictureFile
End If

SystemParametersInfo SPI_SETDESKWALLPAPER, 0, sFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE

bDone = True

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If lblInfo.Tag <> "3" Then
    lblInfo.Caption = "Clicking this with make the current picture your background."
    lblInfo.Tag = 3
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim iPosition As Integer

bDone = False

sPictureFile = tPictureFiles(frmScroller.lPictureIndex).FileName

txtFileName.Text = Right$(sPictureFile, Len(sPictureFile) - LastSlash(sPictureFile))

iPosition = InStr(txtFileName.Text, ".")

txtFileName.Text = Left$(txtFileName.Text, iPosition - 1)

Me.Show

AlwaysOnTop Me.hwnd

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub optLocation_Click(Index As Integer)

On Error GoTo ErrorHandler

If Index = 0 Then
    txtFileName.Enabled = True
Else: txtFileName.Enabled = False
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub optLocation_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If lblInfo.Tag <> CStr(Index) Then
    If Index = 0 Then
        lblInfo.Caption = "With this option selected, Picture Scroller will copy this picture into your Windows directory and use it as your background.  Select this option if your picture is on a remote drive, e.g. a CD."
    Else: lblInfo.Caption = "When this option is selected, Picture Scroller will use this picture as your background regardless of were it is.  Don't use this option if the picture is on a remote drive, e.g. a CD."
    End If

    lblInfo.Tag = Index
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrFlash_Timer()

On Error GoTo ErrorHandler

Static bCount As Byte

If bCount < 7 Then
    If bCount Mod 2 <> 0 Then
        txtFileName.BackColor = vbRed
    Else: txtFileName.BackColor = vbWindowBackground
    End If

    bCount = bCount + 1
Else
    bCount = 0
    tmrFlash.Enabled = False
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub txtFileName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If lblInfo.Tag <> "2" Then
    lblInfo.Caption = "Enter the new file name for the copy of this picture here.  Typically, you don't need to change this."
    lblInfo.Tag = 2
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
