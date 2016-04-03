VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInsertPic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Pictures"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmInsertPic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   5040
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   6
      Top             =   3045
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Timer tmrLoadPreview 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   4200
      Top             =   3000
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1500
      Left            =   4140
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   1500
   End
   Begin MSComctlLib.ListView lstPictures 
      Height          =   4005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2258
      TabIndex        =   3
      Top             =   4095
      Width           =   1695
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   435
      Left            =   158
      TabIndex        =   2
      Top             =   4095
      Width           =   1800
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Select a picture to see a preview of it here:"
      Height          =   585
      Left            =   4200
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInsertPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bChanged As Byte
Private Sub cmdCancel_Click()

Unload Me

End Sub
Private Sub cmdDone_Click()

On Error GoTo ErrorHandler

Dim nIndex As Long
Dim tNewPics() As SELECTED_PIC

ReDim tNewPics(1 To 1)

For nIndex = 1 To tPictureFiles.Count
    If lstPictures.ListItems(nIndex).Checked = True Then
        If UBound(tNewPics) > 100 Then
            MsgBox "A maximum of 100 pictures can be printed a one time.  Please limit your selection down to this number.", vbInformation
            Exit Sub
        Else
            If tNewPics(1).lPictureIndex <> 0 Then ReDim Preserve tNewPics(1 To UBound(tNewPics) + 1)

            tNewPics(UBound(tNewPics)).iValueX = 1
            tNewPics(UBound(tNewPics)).iValueY = 1

            tNewPics(UBound(tNewPics)).lPictureIndex = nIndex
        End If
    End If
Next nIndex

If tNewPics(1).lPictureIndex <> 0 Then
    tSelectedPics = tNewPics

    bChanged = True
Else: ReDim tSelectedPics(0)
End If

Unload Me

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdHelp_Click()

On Error GoTo ErrorHandler

ShowHelp Me.hwnd, "Options_Preview.htm", True

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim nIndex As Long

bChanged = False

If tPictureFiles.Count <> 0 Then
    lstPictures.Checkboxes = True

    For nIndex = 1 To tPictureFiles.Count
        lstPictures.ListItems.Add nIndex, , tPictureFiles(nIndex).FileName
    Next nIndex

    For nIndex = 1 To UBound(tSelectedPics)
        lstPictures.ListItems(tSelectedPics(nIndex).lPictureIndex).Checked = True
    Next nIndex
Else: lstPictures.ListItems.Add , , "No Pictures Added"
End If

lstPictures.ColumnHeaders.Item(1).Width = lstPictures.Width - 50

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lstPictures_ItemClick(ByVal Item As MSComctlLib.ListItem)

tmrLoadPreview.Enabled = False
tmrLoadPreview.Enabled = True

End Sub
Private Sub tmrLoadPreview_Timer()

On Error GoTo ErrorHandler

Dim rPreview As RECT

tmrLoadPreview.Enabled = False

Screen.MousePointer = 11

' Attempt to load the picture.
If GetPicFromIndex(picHidden, lstPictures.SelectedItem.Index, picPreview.Width, picPreview.Height, True) = False Then
    With rPreview
        .Right = picPreview.Width
        .Bottom = picPreview.Height
    End With

    picPreview.Picture = LoadPicture()
    DrawTextAPI picPreview.hdc, "Invalid Picture", 15, rPreview, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP
End If

picPreview = LoadPicture()
BitBlt picPreview.hdc, (picPreview.Width - picHidden.Width) / 2, (picPreview.Height - picHidden.Height) / 2, picHidden.Width, picHidden.Height, picHidden.hdc, 0, 0, vbSrcCopy
picPreview.Refresh

Screen.MousePointer = 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
