VERSION 5.00
Begin VB.Form frmCollection 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2325
   End
End
Attribute VB_Name = "frmCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub DoList(ByVal bWhat As Byte)

Dim nIndex As Long

List1.Clear

If bWhat = 1 Then
    For nIndex = 1 To tPictureFiles.Count
        List1.AddItem tPictureFiles(nIndex).FileName
    Next nIndex
ElseIf bWhat = 2 Then
    For nIndex = 1 To tAddedFolders.Count
        List1.AddItem tAddedFolders(nIndex).FolderName
    Next nIndex
End If

End Sub
