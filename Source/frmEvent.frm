VERSION 5.00
Begin VB.Form frmEvent 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmEvent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent
Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

On Error GoTo ErrorHandler

Dim nIndex As Byte

For nIndex = 0 To 2
    If eventid = hEvent(nIndex) Then
        If nIndex = 2 Then
            BackMusic.MusicEvent eventid
        Else: DirectSound.DirectSoundEvent eventid
        End If

        Exit For
    End If
Next nIndex

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
