VERSION 5.00
Begin VB.Form frmMenu 
   ClientHeight    =   60
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2430
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   60
   ScaleWidth      =   2430
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuAddItem 
         Caption         =   "&Directory"
         Index           =   0
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "&File(s)"
         Index           =   1
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "&Load Saved List"
         Index           =   3
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "&Save Current List"
         Index           =   4
      End
   End
   Begin VB.Menu mnuStretch 
      Caption         =   "MenuStretch"
      Begin VB.Menu mnuStretchItem 
         Caption         =   "&Stretch"
         Index           =   0
      End
      Begin VB.Menu mnuStretchItem 
         Caption         =   "&Normal"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

If bRunMode = RM_SAVER_CONFIG Then
    ' Don't allow the user to save the list.
    frmMenu.mnuAddItem(4).Visible = False
End If

' Subclass the menu so that we know when the
' user holds the mouse over an item.
lPrevWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MenuProc)

End Sub
Private Sub Form_Unload(Cancel As Integer)

' Kill subclassing when we're done.
SetWindowLong Me.hwnd, GWL_WNDPROC, lPrevWndProc

End Sub
Private Sub mnuAddItem_Click(Index As Integer)

frmMain.AddMenuClick Index

End Sub
Private Sub mnuStretchItem_Click(Index As Integer)

frmMain.StretchMenuClick Index

End Sub
