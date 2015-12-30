VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   1830
   ClientTop       =   1890
   ClientWidth     =   7710
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   86
      Top             =   0
      Width           =   7710
   End
   Begin VB.PictureBox picSlider 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2520
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   82
      Top             =   3780
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2520
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1590
      Left            =   735
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.PictureBox picInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   615
      Left            =   0
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   5595
      Width           =   7710
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   2
      Left            =   3084
      TabIndex        =   2
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.Timer tmrLoadPreview 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   4095
         Top             =   105
      End
      Begin VB.TextBox txtLocation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   3990
         WhatsThisHelpID =   77
         Width           =   2910
      End
      Begin MSComctlLib.ImageList imglstIcons 
         Left            =   3720
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   -2147483643
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":059C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstDirectory 
         Left            =   3720
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":06F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0808
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":091A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0A2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0B3E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   2205
         Left            =   210
         TabIndex        =   6
         Top             =   1680
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imglstIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbDirectory 
         Height          =   600
         Left            =   210
         TabIndex        =   9
         Top             =   1050
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   1058
         ButtonWidth     =   1191
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         ImageList       =   "imglstDirectory"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "A&dd"
               Key             =   "ADD"
               Object.ToolTipText     =   "Add a directory or file(s)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "D&elete"
               Key             =   "REMOVE"
               Object.ToolTipText     =   "Remove selected file or whole directory"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   40
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "D&own"
               Key             =   "MOVE_DOWN"
               Object.ToolTipText     =   "Move the selection down one space"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Up"
               Key             =   "MOVE_UP"
               Object.ToolTipText     =   "Move the selection up one space"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   40
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&View"
               Key             =   "PREVIEW"
               Object.ToolTipText     =   "Show a preview of the picture"
               ImageIndex      =   5
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstTemp 
         Height          =   2205
         Left            =   210
         TabIndex        =   96
         Top             =   1680
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "What's Shown:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   345
         TabIndex        =   70
         Top             =   4035
         Width           =   1080
      End
      Begin VB.Label lblInstructions 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":0C50
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   420
         Width           =   4290
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Selected Directories:"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2205
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   3
      Left            =   3084
      TabIndex        =   3
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4110
         Left            =   120
         MousePointer    =   99  'Custom
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   291
         TabIndex        =   80
         Top             =   105
         Width           =   4425
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print Pictures"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   105
            TabIndex        =   85
            Top             =   3465
            Width           =   1620
         End
         Begin VB.VScrollBar VScroll 
            Height          =   975
            LargeChange     =   16
            Left            =   3960
            TabIndex        =   84
            Top             =   120
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   4
      Left            =   3120
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   3
         Left            =   165
         TabIndex        =   16
         Top             =   795
         Visible         =   0   'False
         Width           =   4365
         Begin VB.CheckBox chkSoundEffects 
            Caption         =   "Check this to enable pro&gram sound effects."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   71
            Top             =   3100
            Width           =   3585
         End
         Begin VB.Frame fraPicture 
            Caption         =   "Picture Size"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   1995
            TabIndex        =   40
            Top             =   105
            Width           =   2190
            Begin VB.ComboBox cmbPictureSize 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmMain.frx":0D05
               Left            =   210
               List            =   "frmMain.frx":0D15
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   735
               Width           =   1800
            End
            Begin VB.Label lblPicSize 
               AutoSize        =   -1  'True
               Caption         =   "Select the scrolling size of the pictures below:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   210
               TabIndex        =   91
               Top             =   315
               Width           =   1755
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraScroll 
            Caption         =   "Scroll Direction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   105
            TabIndex        =   36
            Top             =   105
            Width           =   1680
            Begin VB.OptionButton optDirection 
               Caption         =   "&Forward"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   210
               MousePointer    =   99  'Custom
               TabIndex        =   37
               Top             =   315
               WhatsThisHelpID =   65
               Width           =   1125
            End
            Begin VB.OptionButton optDirection 
               Caption         =   "Backwar&d"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   210
               MousePointer    =   99  'Custom
               TabIndex        =   38
               Top             =   600
               WhatsThisHelpID =   66
               Width           =   1125
            End
            Begin VB.OptionButton optDirection 
               Caption         =   "&Random"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   210
               MousePointer    =   99  'Custom
               TabIndex        =   39
               Top             =   885
               WhatsThisHelpID =   67
               Width           =   1125
            End
         End
         Begin VB.Frame fraMusic 
            Caption         =   "Background Music"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   105
            TabIndex        =   19
            Top             =   1365
            Width           =   4080
            Begin VB.CheckBox chkMusicType 
               Caption         =   "A&udio CD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   315
               TabIndex        =   22
               Top             =   315
               WhatsThisHelpID =   70
               Width           =   1215
            End
            Begin VB.CheckBox chkMusicType 
               Caption         =   "Musi&c File"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1680
               TabIndex        =   21
               Top             =   315
               WhatsThisHelpID =   71
               Width           =   1215
            End
            Begin VB.CheckBox chkMusicType 
               Caption         =   "Non&e"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   3045
               TabIndex        =   20
               Top             =   315
               WhatsThisHelpID =   72
               Width           =   795
            End
            Begin VB.PictureBox picMusicType 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   840
               Index           =   2
               Left            =   210
               ScaleHeight     =   810
               ScaleWidth      =   3645
               TabIndex        =   25
               Top             =   630
               Visible         =   0   'False
               Width           =   3675
               Begin VB.Frame fraNormal 
                  BorderStyle     =   0  'None
                  Height          =   810
                  Left            =   0
                  TabIndex        =   31
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   3640
                  Begin VB.ComboBox cmbTracks 
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   1680
                     Style           =   2  'Dropdown List
                     TabIndex        =   95
                     Top             =   360
                     WhatsThisHelpID =   74
                     Width           =   1740
                  End
                  Begin VB.OptionButton optPlayTrack 
                     Caption         =   "Se&lected"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   1
                     Left            =   240
                     TabIndex        =   94
                     Top             =   397
                     WhatsThisHelpID =   74
                     Width           =   1155
                  End
                  Begin VB.OptionButton optPlayTrack 
                     Caption         =   "All Trac&ks"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   0
                     Left            =   240
                     TabIndex        =   93
                     Top             =   120
                     WhatsThisHelpID =   73
                     Width           =   1155
                  End
               End
               Begin VB.Frame fraError 
                  BorderStyle     =   0  'None
                  Height          =   810
                  Left            =   0
                  TabIndex        =   32
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   3640
                  Begin VB.CommandButton cmdRetry 
                     Caption         =   "Retr&y"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   420
                     TabIndex        =   33
                     Top             =   420
                     Width           =   1290
                  End
                  Begin VB.CommandButton cmdCancel 
                     Caption         =   "Cance&l"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   1995
                     TabIndex        =   34
                     Top             =   420
                     Width           =   1290
                  End
                  Begin VB.Label lblCaption 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     Caption         =   "Please insert an audio CD."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   5
                     Left            =   75
                     TabIndex        =   35
                     Top             =   105
                     Width           =   3375
                  End
               End
            End
            Begin VB.PictureBox picMusicType 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   840
               Index           =   1
               Left            =   210
               ScaleHeight     =   810
               ScaleWidth      =   3645
               TabIndex        =   24
               Top             =   630
               Visible         =   0   'False
               Width           =   3675
               Begin VB.OptionButton optLoopMusic 
                  Caption         =   "&Loop Music"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   525
                  TabIndex        =   27
                  Top             =   105
                  WhatsThisHelpID =   75
                  Width           =   1185
               End
               Begin VB.OptionButton optLoopMusic 
                  Caption         =   "Pla&y Once"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   2100
                  TabIndex        =   28
                  Top             =   105
                  WhatsThisHelpID =   76
                  Width           =   1185
               End
               Begin VB.CommandButton cmdBrowse 
                  Caption         =   "Bro&wse"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   2415
                  TabIndex        =   30
                  Top             =   420
                  WhatsThisHelpID =   78
                  Width           =   1065
               End
               Begin VB.TextBox txtMusicFile 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   105
                  TabIndex        =   29
                  Top             =   420
                  WhatsThisHelpID =   77
                  Width           =   2205
               End
            End
            Begin VB.PictureBox picMusicType 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   840
               Index           =   0
               Left            =   210
               ScaleHeight     =   810
               ScaleWidth      =   3645
               TabIndex        =   23
               Top             =   630
               Visible         =   0   'False
               Width           =   3675
               Begin VB.Label lblCaption 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "No Background Music"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   4
                  Left            =   25
                  TabIndex        =   26
                  Top             =   210
                  Width           =   3600
               End
            End
         End
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   1
         Left            =   165
         TabIndex        =   14
         Top             =   795
         Visible         =   0   'False
         Width           =   4365
         Begin VB.TextBox txtTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   83
            Top             =   1950
            Width           =   435
         End
         Begin VB.PictureBox picDial 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            DrawWidth       =   3
            Height          =   2500
            Left            =   1680
            ScaleHeight     =   167
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   167
            TabIndex        =   77
            Top             =   840
            Width           =   2500
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Interval:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   0
            TabIndex        =   79
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "sec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1365
            TabIndex        =   78
            Top             =   1995
            Width           =   255
         End
         Begin VB.Label lblInstructions 
            AutoSize        =   -1  'True
            Caption         =   "Use the dial below to select the interval between pictures."
            Height          =   480
            Index           =   4
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   3855
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   2
         Left            =   165
         TabIndex        =   15
         Top             =   795
         Visible         =   0   'False
         Width           =   4365
         Begin VB.Frame fraMoveDown 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2425
            TabIndex        =   57
            Top             =   2710
            Width           =   1800
            Begin MSComctlLib.Toolbar tlbMoveDown 
               Height          =   270
               Left            =   0
               TabIndex        =   58
               Top             =   -15
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   476
               ButtonWidth     =   3175
               ButtonHeight    =   370
               ImageList       =   "imglstUpDown"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     ImageIndex      =   2
                  EndProperty
               EndProperty
            End
         End
         Begin MSComctlLib.ImageList imglstUpDown 
            Left            =   3570
            Top             =   1155
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   113
            ImageHeight     =   8
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":0D45
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":0FB7
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Frame fraMoveUp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2425
            TabIndex        =   55
            Top             =   715
            Width           =   1800
            Begin MSComctlLib.Toolbar tlbMoveUp 
               Height          =   270
               Left            =   0
               TabIndex        =   56
               Top             =   -15
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   476
               ButtonWidth     =   3175
               ButtonHeight    =   370
               ImageList       =   "imglstUpDown"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     ImageIndex      =   1
                  EndProperty
               EndProperty
            End
         End
         Begin VB.ListBox lstAvailable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            ItemData        =   "frmMain.frx":1229
            Left            =   0
            List            =   "frmMain.frx":122B
            TabIndex        =   53
            Top             =   735
            WhatsThisHelpID =   56
            Width           =   1695
         End
         Begin VB.ListBox lstActual 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            ItemData        =   "frmMain.frx":122D
            Left            =   2410
            List            =   "frmMain.frx":122F
            TabIndex        =   52
            Top             =   1040
            WhatsThisHelpID =   61
            Width           =   1830
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1785
            TabIndex        =   50
            Top             =   1365
            WhatsThisHelpID =   58
            Width           =   540
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1785
            TabIndex        =   49
            Top             =   1890
            WhatsThisHelpID =   59
            Width           =   540
         End
         Begin VB.CheckBox chkRandom 
            Caption         =   "Transitions &Appear Randomly"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   51
            Top             =   3045
            WhatsThisHelpID =   64
            Width           =   2415
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "<<<"
            Height          =   390
            Left            =   1785
            TabIndex        =   48
            Top             =   2550
            WhatsThisHelpID =   60
            Width           =   540
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   ">>>"
            Height          =   390
            Left            =   1785
            TabIndex        =   47
            Top             =   735
            WhatsThisHelpID =   57
            Width           =   540
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Transitions Used:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2415
            TabIndex        =   60
            Top             =   525
            Width           =   1230
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Available:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   59
            Top             =   525
            Width           =   690
         End
         Begin VB.Label lblInstructions 
            AutoSize        =   -1  'True
            Caption         =   "Set the transitions you want to use when scrolling pictures."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   54
            Top             =   105
            Width           =   4125
         End
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   4
         Left            =   165
         TabIndex        =   17
         Top             =   795
         Visible         =   0   'False
         Width           =   4365
         Begin VB.Frame fraColors 
            Caption         =   "Scrolling Colors"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   105
            TabIndex        =   72
            Top             =   2310
            Width           =   4110
            Begin VB.PictureBox picBackColor 
               Height          =   225
               Left            =   1785
               ScaleHeight     =   165
               ScaleWidth      =   1320
               TabIndex        =   74
               Top             =   315
               Width           =   1380
            End
            Begin VB.PictureBox picInfoColor 
               Height          =   225
               Left            =   1785
               ScaleHeight     =   165
               ScaleWidth      =   1320
               TabIndex        =   73
               Top             =   630
               Width           =   1380
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               Caption         =   "Text:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   1260
               TabIndex        =   76
               Top             =   630
               Width           =   360
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               Caption         =   "Background:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   735
               TabIndex        =   75
               Top             =   315
               Width           =   915
            End
         End
         Begin VB.Frame fraScreenSize 
            Caption         =   "Screen Size && Color Depth"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   105
            TabIndex        =   41
            Top             =   105
            Width           =   4095
            Begin VB.OptionButton optScreenSize 
               Caption         =   "Optimize for &Quality"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   210
               TabIndex        =   44
               Top             =   945
               Width           =   2640
            End
            Begin VB.OptionButton optScreenSize 
               Caption         =   "Optimize for Spee&d"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   210
               TabIndex        =   43
               Top             =   630
               Width           =   2640
            End
            Begin VB.ComboBox cmbDisplayModes 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   630
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   1575
               Width           =   2745
            End
            Begin VB.OptionButton optScreenSize 
               Caption         =   "&Custom"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   210
               TabIndex        =   45
               Top             =   1260
               Width           =   2640
            End
            Begin VB.OptionButton optScreenSize 
               Caption         =   "&Use the Current Screen Settings"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   210
               TabIndex        =   42
               Top             =   315
               Width           =   2640
            End
         End
      End
      Begin MSComctlLib.TabStrip tbsOptions 
         Height          =   3840
         Left            =   105
         TabIndex        =   12
         Top             =   420
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   6773
         MultiRow        =   -1  'True
         Style           =   2
         HotTracking     =   -1  'True
         TabMinWidth     =   1623
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "I&nterval"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Transitions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Miscellaneous"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ad&vanced"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblInstructions 
         AutoSize        =   -1  'True
         Caption         =   "Click a button to set related options:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   13
         Top             =   105
         Width           =   2535
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   1
      Left            =   3084
      TabIndex        =   1
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.Label lblInstructions 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "* Click Logo for Credits *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   81
         Top             =   3120
         Width           =   4395
      End
      Begin VB.Label lblInstructions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Move your mouse over just about anything to see a description below. ¬"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1470
         TabIndex        =   63
         Top             =   3885
         Width           =   3030
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInstructions 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   62
         Top             =   840
         Width           =   4125
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Welcome to Picture Scroller!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   135
         TabIndex        =   61
         Top             =   210
         Width           =   4335
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   5
      Left            =   3150
      TabIndex        =   5
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.Timer tmrFlash 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   105
         Top             =   3255
      End
      Begin VB.Label lblInstructions 
         AutoSize        =   -1  'True
         Caption         =   "Press ""Enter"" to begin."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   92
         Top             =   3990
         Width           =   1620
      End
      Begin VB.Label lblNoPics 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "No pictures have been added"
         Height          =   240
         Left            =   735
         TabIndex        =   89
         Top             =   3255
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "# of Pictures:"
         Height          =   240
         Index           =   12
         Left            =   210
         TabIndex        =   88
         Top             =   1365
         Width           =   1350
      End
      Begin VB.Label lblNumOfPictures 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   2100
         TabIndex        =   87
         Top             =   1365
         Width           =   75
      End
      Begin VB.Label lblBeginScrolling 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "<< Begin >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1410
         TabIndex        =   68
         Top             =   2415
         Width           =   1740
      End
      Begin VB.Label lblScrollDirection 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   2100
         TabIndex        =   67
         Top             =   945
         Width           =   75
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   2100
         TabIndex        =   66
         Top             =   525
         Width           =   75
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Direction:"
         Height          =   240
         Index           =   9
         Left            =   210
         TabIndex        =   65
         Top             =   945
         Width           =   1665
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Interval:"
         Height          =   240
         Index           =   8
         Left            =   210
         TabIndex        =   64
         Top             =   525
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COLOR_BTNFACE = 15
Const COLOR_HIGHLIGHT = 13

Const COLOR_BACK_REG = &HAC7556
Const COLOR_LINES = vbBlack
Const COLOR_TEXT_REG = &HFFFF00
Const COLOR_TEXT_OVER = &HCCFF0

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long

' ------------------------------------
' Used to get a folder from the user and to add a
' file to the recent docs list.

Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type ITEMIDLIST
    mkid As Long
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As ITEMIDLIST
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
' ------------------------------------
' Used for searching for pictures

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
' ------------------------------------
' These are used to get the system icon of
' different types of picture files.

Const SHGFI_ICON = &H100
Const SHGFI_SMALLICON = &H1
Const SHGFI_USEFILEATTRIBUTES = &H10
Const DI_NORMAL = &H3

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
' ------------------------------------
' Used to find the menu item that's selected.

Const MF_BYPOSITION = &H400&
Const MF_HILITE = &H80&

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
' ------------------------------------
' Used for drawing the interval dial and previewing pictures.

Const DIR_UP = 0
Const DIR_DOWN = 1
Const NUM_OF_PREVIEWS = 16

Dim bDown As Byte
Dim iSize As Integer
Dim iPoint As Integer
Dim iYPos As Integer
Dim lPreviewPics(1 To NUM_OF_PREVIEWS) As Long
Dim bNextFree As Byte

Dim lLastVal As Long

Dim pClickPos As POINTAPI
' ------------------------------------

' Hold which button the mouse is over
Dim bMouseOver As Byte
' Holds which button was last clicked
Dim bLastClicked As Byte

' Holds the current tab of the options section
Dim bCurrentOption As Byte

' The currently displayed folder
Dim sCurrentFolder As String

' Set to true so that the checkboxes' click events
' won't do anything just because we set their value
' through code
Dim bByCode As Byte

Dim bUnloadNow As Byte

' Holds the dimensions for each button
Dim rButtons() As RECT
' Specifies the number of buttons
Dim bNumOfOptions As Byte

Dim bFlashCount As Byte

Dim iCirclePoints(19) As POINTAPI
Private Sub chkMusicType_Click(Index As Integer)

On Error GoTo ErrorHandler

Dim nIndex As Byte

If bByCode = True Then Exit Sub

SetMusicType Index

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub chkMusicType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 364 + Index

End Sub
Private Sub chkRandom_Click()

On Error GoTo ErrorHandler

tProgramOptions.bRandomTransitions = chkRandom.Value

UpdateTransButtons

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub chkRandom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 348

End Sub
Private Sub chkSoundEffects_Click()

tProgramOptions.bSoundEffects = chkSoundEffects.Value

End Sub
Private Sub chkSoundEffects_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 375

End Sub
Private Sub cmbDisplayModes_Click()

On Error GoTo ErrorHandler

With tProgramOptions.tDisplayMode
    .iWidth = DirectDraw.ModeWidth(cmbDisplayModes.ListIndex + 1)
    .iHeight = DirectDraw.ModeHeight(cmbDisplayModes.ListIndex + 1)
    .bBPP = DirectDraw.ModeBPP(cmbDisplayModes.ListIndex + 1)
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmbDisplayModes_GotFocus()

optScreenSize(3).Value = True

End Sub
Private Sub cmbPictureSize_Click()

tProgramOptions.bPictureSize = cmbPictureSize.ListIndex

End Sub
Private Sub cmbTracks_Click()

tProgramOptions.iTrackNumber = cmbTracks.ListIndex + 1

End Sub
Private Sub cmdAdd_Click()

On Error GoTo ErrorHandler

Dim iIndexValue As Integer

If lstAvailable.ListIndex = -1 Then Exit Sub

iIndexValue = lstAvailable.ListIndex

lstActual.AddItem lstAvailable.List(iIndexValue)
lstActual.ItemData(lstActual.ListCount - 1) = lstAvailable.ItemData(iIndexValue)
lstAvailable.RemoveItem iIndexValue

If iIndexValue < lstAvailable.ListCount Then
    lstAvailable.ListIndex = iIndexValue
Else: lstAvailable.ListIndex = lstAvailable.ListCount - 1
End If

UpdateTransButtons
UpdateTransitions

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 342

End Sub
Private Sub cmdAddAll_Click()

On Error GoTo ErrorHandler

lstAvailable.ListIndex = 0

Do Until lstAvailable.ListCount = 0
    cmdAdd_Click
Loop

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdAddAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 341

End Sub
Private Sub cmdBrowse_Click()

On Error GoTo ErrorHandler

Dim sSelectedFiles() As String
Dim sFolder As String

' Show the OpenFile CommonDialog box
sFolder = OpenFileDialog(Me.hwnd, MUSIC, sSelectedFiles)

' If the user didn't select a file, exit
If sSelectedFiles(0) = "" Then Exit Sub

tProgramOptions.sMusicFile = sSelectedFiles(0)
txtMusicFile.Text = sSelectedFiles(0)
txtMusicFile.SetFocus

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 374

End Sub
Private Sub cmdCancel_Click()

On Error GoTo ErrorHandler

Dim nIndex As Byte
Dim bLastFound As Byte

' Loop through the options and find the option the
' user clicked last.
For nIndex = 0 To 2
    If chkMusicType(nIndex).Tag = "LAST" Then
        ' Set to true so we know we've found it.
        bLastFound = True

        SetMusicType nIndex
        Exit For
    End If
Next nIndex

' If the user hasn't clicked an option before,
' set it to none.
If bLastFound = False Then SetMusicType 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 370

End Sub
Private Sub cmdPrint_Click()

On Error GoTo ErrorHandler

Me.Enabled = False

frmPrint.Show , Me
DoEvents

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 200

End Sub
Private Sub cmdRemove_Click()

On Error GoTo ErrorHandler

Dim iIndexValue As Integer

iIndexValue = lstActual.ListIndex

lstAvailable.AddItem lstActual.List(iIndexValue)
lstAvailable.ItemData(lstAvailable.ListCount - 1) = lstActual.ItemData(iIndexValue)
lstActual.RemoveItem iIndexValue

If iIndexValue < lstActual.ListCount Then
    lstActual.ListIndex = iIndexValue
Else: lstActual.ListIndex = lstActual.ListCount - 1
End If

UpdateTransButtons
UpdateTransitions

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 343

End Sub
Private Sub cmdRemoveAll_Click()

On Error GoTo ErrorHandler

lstActual.ListIndex = 0

Do Until lstActual.ListCount = 0
    cmdRemove_Click
Loop

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub cmdRemoveAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 344

End Sub
Private Sub cmdRetry_Click()

FillTracks

End Sub
Private Sub cmdRetry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 369

End Sub
Private Sub Form_Click()

On Error GoTo ErrorHandler

Dim hBrush As Long

If bMouseOver > 0 Then
    ' Play "click" sound if sound effects enabled.
    If tProgramOptions.bSoundEffects = 1 And DirectSound.bInitOK = True Then DirectSound.PlaySound "SELECT", False

    ' Reset the previous option clicked
    If bMouseOver <> 6 And bLastClicked > 0 And bLastClicked <> bMouseOver Then
        ' Refill the last option with the back color
        hBrush = CreateSolidBrush(COLOR_BACK_REG)
        FillRect Me.hdc, rButtons(bLastClicked), hBrush

        ' Draw the button's caption in the non-selected
        ' state
        Me.ForeColor = COLOR_TEXT_REG
        DrawTextAPI Me.hdc, LoadResString(bLastClicked), Len(LoadResString(bLastClicked)), rButtons(bLastClicked), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

        DeleteObject hBrush

        ' Hide the previous frame
        fraSteps(bLastClicked).Visible = False
    End If

    If bMouseOver = 3 Then
        lLastVal = 0

        picShow.Cls
        picShow.Refresh
        picSlider.Cls
        picSlider.Refresh

        If tPictureFiles.Count <= 4 Then
            VScroll.Visible = False
            iSize = (picShow.ScaleHeight - cmdPrint.Height - 30) / 2
            picSlider.Width = iSize * 2 + 30
            picSlider.Height = (iSize * (NUM_OF_PREVIEWS \ 2)) + ((NUM_OF_PREVIEWS \ 2) * 10) + 10
        Else
            bByCode = 1

            If tPictureFiles.Count Mod 2 = 0 Then
                VScroll.Max = tPictureFiles.Count / 2
            Else
                VScroll.Max = (tPictureFiles.Count + 1) / 2
            End If

            VScroll.Max = VScroll.Max - 3

            iSize = (picShow.ScaleHeight - cmdPrint.Height - 30) / 2
            picSlider.Width = iSize * 2 + 30 + VScroll.Width
            picSlider.Height = (iSize * (NUM_OF_PREVIEWS \ 2)) + ((NUM_OF_PREVIEWS \ 2) * 10) + 10

            If VScroll.Max * 2 + 1 < tPictureFiles.Count Then VScroll.Max = VScroll.Max + 1

            bByCode = 0

            VScroll.Visible = True
        End If

        DoEvents
    End If

    If bMouseOver <> 6 Then
        ' Set which button is selected
        bLastClicked = bMouseOver

        ' First setup the next frame to be shown
        fraSteps(bLastClicked).Visible = True

        ' Draw the background for the button
        hBrush = GetSysColorBrush(COLOR_BTNFACE)
        FillRect Me.hdc, rButtons(bLastClicked), hBrush

        ' Draw the button's caption
        Me.ForeColor = GetSysColor(COLOR_HIGHLIGHT)
        DrawTextAPI Me.hdc, LoadResString(bLastClicked), Len(LoadResString(bLastClicked)), rButtons(bLastClicked), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

        SetupOptions bLastClicked
    Else: SetupOptions bMouseOver
    End If

    Me.Refresh

    DoEvents
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

Dim bClickButton As Byte

' Check the user's key stroke and do as told.

If bLastClicked = 5 And KeyCode = vbKeyReturn Then
    ' If we're on the "Begin Scrolling" selection
    ' then allow the user to press "Enter" to begin.
    lblBeginScrolling_Click
ElseIf KeyCode = vbKeyF1 Then
    ContextHelp False
ElseIf Shift And vbAltMask Then
    ' Alt keys.
    If bLastClicked = 2 Then
        ' If the selected button is the "Select
        ' Picture(s)", then allow them to use the
        ' ALT plus a key to select the toolbar.
        Select Case KeyCode
            Case vbKeyD: tlbDirectory_ButtonClick tlbDirectory.buttons(1): Exit Sub
            Case vbKeyE: tlbDirectory_ButtonClick tlbDirectory.buttons(2): Exit Sub
            Case vbKeyO: tlbDirectory_ButtonClick tlbDirectory.buttons(4): Exit Sub
            Case vbKeyU: tlbDirectory_ButtonClick tlbDirectory.buttons(5): Exit Sub
            Case vbKeyV
                tlbDirectory.buttons(7).Value = Abs(tlbDirectory.buttons(7).Value - 1)
                tlbDirectory_ButtonClick tlbDirectory.buttons(7)

                Exit Sub
        End Select
    End If

    ' If we're here then the user hasn't pressed
    ' one of the previously checked keys.

    ' Allow the user to press ALT plus a key to
    ' access the various button options on the left.
    Select Case KeyCode
        Case vbKeyI: bClickButton = 1
        Case vbKeyS: bClickButton = 2
        Case vbKeyP: bClickButton = 3
        Case vbKeyA: bClickButton = 4
        Case vbKeyB: bClickButton = 5
        Case vbKeyH: bClickButton = 6
    End Select

    If bClickButton <> 0 And bClickButton <> bLastClicked Then
        bMouseOver = bClickButton
        Form_Click
    End If
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler

Dim rOptions As RECT
Dim hBrush As Long
Dim nIndex As Byte

With Me
    .Caption = App.Title
    .Left = (Screen.Width - .Width) \ 2
    .Top = (Screen.Height - .Height) \ 2
End With

WaitProcess "Loading Settings", False, Me

' ----------------------------------------------
' Main Form
If bRunMode = RM_SAVER_CONFIG Then
    ' Set the caption
    Me.Caption = Me.Caption & " - Screen Saver Configuration"

    ' Instructions tab
    lblInstructions(0).Caption = LoadResString(43)

    ' Bottom welcome text
    DrawTip 41
Else
    lblInstructions(0).Caption = LoadResString(42)

    DrawTip 40
End If

DoEvents

' ----------------------------------------------
' Logo

With picHidden
    .Width = 514
    .Height = 81
    .Picture = PictureFromBits(LoadResData(12, "PICS"))
End With

SetStretchBltMode picLogo.hdc, STRETCH_DELETESCANS
StretchBlt picLogo.hdc, 0, 0, picLogo.Width, picLogo.Height, picHidden.hdc, 0, 0, picHidden.Width, picHidden.Height, vbSrcCopy

picHidden.Picture = LoadPicture()

DoEvents

' ----------------------------------------------
' Control sizes

lstTemp.ListItems.Add , , "Loading..."

cmdPrint.Left = 0
cmdPrint.Top = picShow.ScaleHeight - cmdPrint.Height
cmdPrint.Width = picShow.ScaleWidth

' Calculate the size for the preview pictures
VScroll.Top = 0
VScroll.Height = picShow.ScaleHeight - cmdPrint.Height
VScroll.Left = picShow.ScaleWidth - VScroll.Width
VScroll.Min = 0

picShow.MouseIcon = LoadResPicture(1, vbResCursor)

' Position the option pages exactly
For nIndex = 1 To 4
    fraOptions(nIndex).Left = tbsOptions.ClientLeft
    fraOptions(nIndex).Top = tbsOptions.ClientTop
    fraOptions(nIndex).Width = tbsOptions.ClientWidth
    fraOptions(nIndex).Height = tbsOptions.ClientHeight
Next nIndex

DoEvents

' ----------------------------------------------
' Picture (document) icons

' Retrieve all the different picture icons
GetPicIcons

DoEvents

' ----------------------------------------------
' Left option area

' Setup the options section area
With rOptions
    .Left = 0
    .Top = picLogo.Height
    .Right = fraSteps(1).Left
    .Bottom = picInfo.Top
End With

' Create a dark blue brush
hBrush = CreateSolidBrush(COLOR_BACK_REG)
' Draw a box for the options on the left side
FillRect Me.hdc, rOptions, hBrush
DeleteObject hBrush

' Find out the number of buttons availables
bNumOfOptions = LoadResString(20)

' Set up the buttons array to the correct number
ReDim rButtons(1 To bNumOfOptions)

DoEvents

' ----------------------------------------------
' Left option buttons

For nIndex = 1 To bNumOfOptions
    ' Resize all the option frames to the correct size
    If nIndex <> bNumOfOptions Then
        With fraSteps(nIndex)
            .Left = Me.ScaleWidth * 0.4
            .Width = Me.ScaleWidth - .Left
            .Top = picLogo.Height
            .Height = picInfo.Top - .Top
        End With
    End If

    ' Draw divider lines and the options' text
    With rButtons(nIndex)
        If nIndex = 1 Then
            .Top = rOptions.Top
        ElseIf nIndex = bNumOfOptions Then
            .Top = picInfo.Top - 29
        Else: .Top = rButtons(nIndex - 1).Bottom + 1
        End If

        .Right = rOptions.Right
        .Bottom = .Top + 25
    End With

    Me.ForeColor = COLOR_LINES

    If nIndex = bNumOfOptions Then
        MoveToEx Me.hdc, 0, rButtons(nIndex).Top, 0
        LineTo Me.hdc, rButtons(nIndex).Right, rButtons(nIndex).Top
    End If

    MoveToEx Me.hdc, 0, rButtons(nIndex).Bottom, 0
    LineTo Me.hdc, rButtons(nIndex).Right, rButtons(nIndex).Bottom

    Me.ForeColor = COLOR_TEXT_REG
    DrawTextAPI Me.hdc, LoadResString(nIndex), Len(LoadResString(nIndex)), rButtons(nIndex), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

    DoEvents
Next nIndex

' ----------------------------------------------
' User options, sound, etc.

' Select the first option
bMouseOver = 1
Form_Click

If bRunMode = RM_SAVER_CONFIG Then
    ' Load the list of ScreenSaver pictures
    LoadSavedList sAppPath & "SSList.pcs", False, False
ElseIf tProgramOptions.bNotFirstStart = False Then
    ' Load the example pictures on the first start.
    LoadSavedList sAppPath & "Examples.pcs", False, True
Else
    ' Since we aren't loading any pictures the
    ' toolbar buttons should get disabled.
    UpdateDirButtons
End If

' Set all the options to the user's last.
SetUserOptions

DoEvents
InitDirectSound
DoEvents

' If we're to play sounds, then play the welcome message.
If tProgramOptions.bSoundEffects = 1 And DirectSound.bInitOK = True Then
    ' Play the "Welcome Words" :)
    DirectSound.PlaySound "WELCOME", False
End If

DoEvents

EndWaitProcess Me

Me.Show
DoEvents

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Dim nIndex As Byte
Dim bInButton As Byte

For nIndex = 1 To bNumOfOptions
    With rButtons(nIndex)
        ' See if the mouse is within the area of this
        ' particular button
        If X > .Left And X < .Right And Y > .Top And Y < .Bottom Then
            ' Just exit if this is the same button we
            ' were on before.
            If nIndex = bMouseOver Then Exit Sub

            If bMouseOver <> 0 And bLastClicked <> bMouseOver Then
                If nIndex = bLastClicked Or nIndex <> bMouseOver Then
                    Me.ForeColor = COLOR_TEXT_REG
                    DrawTextAPI Me.hdc, LoadResString(bMouseOver), Len(LoadResString(bMouseOver)), rButtons(bMouseOver), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

                    If bMouseOver <> nIndex Then bMouseOver = nIndex
                End If
            End If

            DrawTip 20 + nIndex

            If nIndex = bLastClicked Then
                Me.Refresh
                Exit Sub
            End If

            ' Set the new last button as the current
            ' button (for the next time to mouse moves)
            If bMouseOver <> nIndex Then bMouseOver = nIndex

            ' Draw the button text in a different color
            Me.ForeColor = COLOR_TEXT_OVER

            DrawTextAPI Me.hdc, LoadResString(nIndex), Len(LoadResString(nIndex)), rButtons(nIndex), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

            bInButton = True

            ' We can now exit; the mouse would only
            ' be on one button
            Exit For
        End If
    End With
Next nIndex

' If the mouse is outside the area of
' any button but a button has been
' previously hovered on, we need to reset
' that button's colors
If bInButton = False And bMouseOver > 0 Then
    If bMouseOver <> bLastClicked Then
        Me.ForeColor = COLOR_TEXT_REG
        DrawTextAPI Me.hdc, LoadResString(bMouseOver), Len(LoadResString(bMouseOver)), rButtons(bMouseOver), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

        ' No button is now selected
        bMouseOver = 0

        ' Clear the picInfo of any information
        picInfo.Cls
    End If
End If

Me.Refresh

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo ErrorHandler

Screen.MousePointer = 11

bUnloadNow = True

' Save the user's options.
SaveUserOptions

If bRunMode = RM_SAVER_CONFIG Then SaveCurrentList sAppPath & "SSList.pcs"

Set DirectSound = Nothing

ShowHelp 0, "<<CLOSEALL>>", False

Screen.MousePointer = 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub fraPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 363

End Sub
Private Sub fraSteps_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If Index = 5 Then
    With lblBeginScrolling
        If .Tag = "OVER" Then
            .FontUnderline = False
            .ForeColor = vbButtonText
            .Tag = ""
        End If
    End With
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lblBeginScrolling_Click()

On Error GoTo ErrorHandler

If tPictureFiles.Count = 0 Then
    ' Tell the user that no pictures have been added.
    lblNoPics.ForeColor = vbButtonText
    lblNoPics.Visible = True

    bFlashCount = 0
    tmrFlash.Enabled = True
    Exit Sub
End If

If tProgramOptions.bNotFirstStart = False Then MsgBox "In order to show/hide the scrolling controls, press the 'H' key." & vbCr & vbCr & "This message will only be shown once.", vbOKOnly + vbInformation

Me.Hide
DoEvents

' Kill DirectSound so that the scrolling form can use it.
Set DirectSound = Nothing

Load frmScroller
DoEvents

Unload frmScroller
DoEvents

Me.Show
DoEvents

InitDirectSound

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lblBeginScrolling_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

With lblBeginScrolling
    If .Tag = "" Then
        .FontUnderline = True
        .ForeColor = vbBlue
        .Tag = "OVER"
    End If
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lblPicSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 363

End Sub
Private Sub lstActual_Click()

UpdateTransButtons

End Sub
Private Sub lstActual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 345

End Sub
Private Sub lstAvailable_Click()
    
UpdateTransButtons

End Sub
Private Sub lstAvailable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 340

End Sub
Private Sub lstFiles_DblClick()

On Error GoTo ErrorHandler

Dim sPath As String

If lstFiles.SelectedItem.Key = "NO_FILES" Then Exit Sub

If lstFiles.SelectedItem.Key = "UP_LEVEL" Then
    AddFilesToList ""
Else
    ' Figure out the path of the selected file
    sPath = GetItemPath(lstFiles.SelectedItem.Index)

    If GetAttr(sPath) And vbDirectory Then AddFilesToList sPath
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lstFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

UpdateDirButtons

' Stop the current delay for loading a picture: we
' just changed to a different picture.
tmrLoadPreview.Enabled = False

' Exit if it's not applicable to load the preview.
If tlbDirectory.buttons("PREVIEW").Value = tbrUnpressed Or _
    Item.Key = "NO_FILES" Or Item.Key = "UP_LEVEL" Or _
    sCurrentFolder = "" Then
        picPreview.Picture = LoadPicture()
Else
    ' Now set the timer to wait a bit before loading
    ' a preview, so that Windows will make this item
    ' selected and so that the user will have a chance
    ' so switch to a differnt picture if they want to.
    tmrLoadPreview.Enabled = True
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lstFiles_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

Dim nIndex As Long

' Allow the user to press Ctrl+A to select all the files
If KeyCode = vbKeyA And (Shift And vbCtrlMask) > 0 Then
    For nIndex = 1 To lstFiles.ListItems.Count
        lstFiles.ListItems(nIndex).Selected = True
    Next nIndex
ElseIf KeyCode = vbKeyDelete Then
    tlbDirectory_ButtonClick tlbDirectory.buttons(2)
ElseIf KeyCode = vbKeyReturn And sCurrentFolder = "" Then
    lstFiles_DblClick
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 100

End Sub
Private Sub optDirection_Click(Index As Integer)

tProgramOptions.bScrollDirection = Index

End Sub
Private Sub optDirection_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 360 + Index

End Sub
Private Sub optLoopMusic_Click(Index As Integer)

tProgramOptions.bLoopMusic = Index

End Sub
Private Sub optLoopMusic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 371 + Index

End Sub
Private Sub optPlayTrack_Click(Index As Integer)

tProgramOptions.bPlayTrack = Index
cmbTracks.Enabled = Index

End Sub
Private Sub optPlayTrack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 367 + Index

End Sub
Private Sub optScreenSize_Click(Index As Integer)

tProgramOptions.bScreenSetting = Index

End Sub
Private Sub optScreenSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 380 + Index

End Sub
Private Sub picBackColor_Click()

On Error GoTo ErrorHandler

Dim lNewColor As Long

lNewColor = SelectColor(Me.hwnd, tProgramOptions.lBackColor)

If lNewColor <> -1 Then
    tProgramOptions.lBackColor = lNewColor
    picBackColor.BackColor = lNewColor
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picBackColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 384

End Sub
Private Sub picDial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    bDown = 1

    picDial_MouseMove Button, 0, X, Y
End If

End Sub
Private Sub picDial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Dim lColor As Long
Dim nIndex As Integer
Dim iShortest As Integer
Dim dTemp As Double
Dim dShortest As Double

If bDown = 1 Then
    lColor = picDial.ForeColor

    picDial.ForeColor = picDial.BackColor

    picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    picDial.ForeColor = lColor

    picDial.PSet (iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    For nIndex = 0 To 19
        dTemp = Sqr((iCirclePoints(nIndex).X - X) ^ 2 + (iCirclePoints(nIndex).Y - Y) ^ 2)

        If dTemp < dShortest Or dShortest = 0 Then
            dShortest = dTemp
            iShortest = nIndex
        End If
    Next nIndex

    iPoint = iShortest

    txtTime.Text = (iPoint + 1)

    picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    tProgramOptions.iInterval = (iPoint + 1) * 1000
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picDial_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

bDown = 0

End Sub
Private Sub picInfoColor_Click()

On Error GoTo ErrorHandler

Dim lNewColor As Long

lNewColor = SelectColor(Me.hwnd, tProgramOptions.lInfoColor)

If lNewColor <> -1 Then
    tProgramOptions.lInfoColor = lNewColor
    picInfoColor.BackColor = lNewColor
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picInfoColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 385

End Sub
Private Sub picLogo_Click()

On Error GoTo ErrorHandler

Me.Hide
DoEvents

Load frmCredits
DoEvents

Unload frmCredits
DoEvents

Me.Show
DoEvents

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 44

End Sub
Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 45

End Sub
Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 101

End Sub
Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

If lSlot <> 0 And Button = 1 Then
    DrawTip 0, "Click picture or press any key to exit."

    Me.Enabled = False

    frmPreview.lPictureIndex = lSlot
    frmPreview.Show , Me
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Static lOldSlot As Long

If Button = 1 Then Exit Sub

If tPictureFiles.Count = 0 Then Exit Sub

lSlot = FindSlot(X, Y, lOldSlot)

If lSlot = 0 Then
    DrawTip 0, " "
ElseIf lPreviewPics(lSlot) <> 0 Then
    DrawTip 0, tPictureFiles(lPreviewPics(lSlot)).FileName
Else: DrawTip 0, " "
End If

lOldSlot = lSlot

If lSlot <> 0 Then lSlot = lPreviewPics(lSlot)

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub picShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bDown = 1 Then
    picShow.MouseIcon = LoadResPicture(1, vbResCursor)

    bDown = 0
End If

End Sub
Private Sub tbsOptions_Click()

On Error GoTo ErrorHandler

Dim bLastOption As Byte

bLastOption = bCurrentOption

' Show the the frame of the option clicked.
bCurrentOption = tbsOptions.SelectedItem.Index
fraOptions(bCurrentOption).Visible = True
DoEvents

' We do this here instead of at load time, because
' we don't want to bother the user until they go here.
If bCurrentOption = 3 Then SetMusicType tProgramOptions.bMusicType

' Hide the previous option frame, unless of course it is us.
If bLastOption <> bCurrentOption Then
    fraOptions(bLastOption).Visible = False
    DoEvents
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tbsOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Dim nIndex As Byte

For nIndex = 1 To tbsOptions.Tabs.Count
    With tbsOptions.Tabs(nIndex)
        If X > .Left And X < .Left + .Width And (Y + tbsOptions.Top) > .Top And (Y + tbsOptions.Top) < .Top + .Height Then
            DrawTip 299 + nIndex
        End If
    End With
Next nIndex

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tlbDirectory_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

Dim iToolbarX As Integer
Dim iToolbarY As Integer

Dim iLastSlash As Integer

If Button.Enabled = False Then Exit Sub

' Make sure the button is refreshed.
DoEvents

' Calculate the position for the menu
iToolbarX = fraSteps(1).Left + (tlbDirectory.Left \ Screen.TwipsPerPixelX) + (Button.Left \ Screen.TwipsPerPixelX) - 1
iToolbarY = fraSteps(1).Top + (tlbDirectory.Top \ Screen.TwipsPerPixelY) + (Button.Height \ Screen.TwipsPerPixelY) + 2

' Show the correct menu for the button clicked
Select Case Button.Key
    Case "ADD"
        ' We don't need to set the state of the "Save List"
        ' option if we're in ScreenSaver Configuration.
        If bRunMode <> RM_SAVER_CONFIG Then
            If lstFiles.ListItems(1).Key = "NO_FILES" Then
                frmMenu.mnuAddItem(4).Enabled = False
            Else: frmMenu.mnuAddItem(4).Enabled = True
            End If

            If tProgramOptions.sRecent1 <> "" Then
                iLastSlash = LastSlash(tProgramOptions.sRecent1)

                frmMenu.mnuAddItem(6).Caption = "1) " & Right$(tProgramOptions.sRecent1, Len(tProgramOptions.sRecent1) - iLastSlash)
                frmMenu.mnuAddItem(6).Tag = tProgramOptions.sRecent1
                frmMenu.mnuAddItem(6).Visible = True

                If tProgramOptions.sRecent2 <> "" Then
                    iLastSlash = LastSlash(tProgramOptions.sRecent2)

                    frmMenu.mnuAddItem(7).Caption = "2) " & Right$(tProgramOptions.sRecent2, Len(tProgramOptions.sRecent2) - iLastSlash)
                    frmMenu.mnuAddItem(7).Tag = tProgramOptions.sRecent2
                    frmMenu.mnuAddItem(7).Visible = True
                End If

                frmMenu.mnuAddItem(5).Visible = True
            Else
                frmMenu.mnuAddItem(5).Visible = False
                frmMenu.mnuAddItem(6).Visible = False
                frmMenu.mnuAddItem(7).Visible = False
            End If

        End If

        ShowPopMenu frmMenu.mnuAdd, iToolbarX, iToolbarY
    Case "REMOVE"
        RemoveItems
    Case "MOVE_UP", "MOVE_DOWN"
        If sCurrentFolder <> "" Then
            MovePicture Button.Key
        Else: MoveFolder Button.Key
        End If

        UpdateDirButtons
    Case "PREVIEW"
        ' Show or hide the preview picturebox depending
        ' on if this button is checked or not
        picPreview.Visible = Button.Value

        ' Preview the currently selected item
        lstFiles_ItemClick lstFiles.SelectedItem
End Select

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub SetupOptions(ByVal bStep As Byte)

' Purpose: Setup various settings depending on the
'   option selected (the buttons on the left).

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim sText As String

Select Case bStep
    Case 2
        ' Add files to the list if necessary
        If lstFiles.ListItems.Count = 0 Then AddFilesToList sCurrentFolder

        ' If the user has previously shown the preview
        ' button, then show it again.
        If tlbDirectory.buttons("PREVIEW").Value = tbrPressed Then picPreview.Visible = True
    Case 3
        DoEvents
        bNextFree = 1
        bByCode = 1
        VScroll.Value = 0
        bByCode = 0
        iYPos = 0

        For nIndex = 1 To UBound(lPreviewPics)
            lPreviewPics(nIndex) = 0
        Next nIndex

        If tPictureFiles.Count = 0 Then
            picShow.Cls

            sText = "No Pictures Added"
            With picShow
                DrawTextAPI .hdc, sText, Len(sText), DirectDraw.GetRect(0, 0, .ScaleWidth, .ScaleHeight), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            End With
        Else
            Screen.MousePointer = 11
            LoadRow DIR_DOWN, 1
            LoadRow DIR_DOWN, 1
            Screen.MousePointer = 0
        End If
    Case 4
        If bCurrentOption = 0 Then
            bCurrentOption = 1
            fraOptions(bCurrentOption).Visible = True
        End If
    Case 5
        lblInterval.Caption = tProgramOptions.iInterval \ 1000 & " seconds"

        Select Case tProgramOptions.bScrollDirection
            Case 0: lblScrollDirection.Caption = "Forwards"
            Case 1: lblScrollDirection.Caption = "Backwards"
            Case 2: lblScrollDirection.Caption = "Randomly"
        End Select

        lblNumOfPictures.Caption = IIf(tPictureFiles.Count = 0, "NONE", tPictureFiles.Count)
    Case 6: ContextHelp True
End Select

' Reset values not needed for the other options.
If bMouseOver <> 2 Then
    If picPreview.Visible = True Then picPreview.Visible = False
End If

If bMouseOver <> 5 Then
    If tmrFlash.Enabled = True Then
        tmrFlash.Enabled = False
        lblNoPics.Visible = False
    End If
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub SearchForPics(ByVal sPath As String, ByVal bRecursive As Byte)

' Purpose: Searchs through the specified directory
'   and adds only picture files to the array
'   "tPictureFiles". This function is called by itself
'   in order to search subfolders.

On Error GoTo ErrorHandler

Dim hFindHandle As Long
Dim lReturnVal As Long
Dim iPosition As Integer
Dim sResult As String
Dim bType As Byte
Dim lFolderID As Long

' Variables for retrieving file information
Dim tFileData As WIN32_FIND_DATA

' Searching for all occurences of the file in the directory
hFindHandle = FindFirstFile(sPath & "*.*", tFileData)

' Make sure we go through the DoLoop, unless our initial
' search result is -1
lReturnVal = IIf(hFindHandle = -1, 0, 1)

' If the user does cancel the operation, then continue
Do While DoEvents And lReturnVal <> 0 And bCancelOp <> True
    ' Get the whole path of the file found
    iPosition = InStr(tFileData.cFileName, vbNullChar)
    If iPosition <> 0 Then
        sResult = Left$(tFileData.cFileName, iPosition - 1)
    Else: sResult = tFileData.cFileName
    End If

    If sResult <> "." And sResult <> ".." Then
        ' If the file isn't actually a folder...
        If (tFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
            ' See if the file is a picture
            bType = ConfirmType(sPath & sResult)

            ' If it is...
            If bType <> 0 Then
                ' If we don't have a folder ID yet
                ' (because this is the first file
                ' found), then get an/the ID of it.
                If lFolderID = 0 Then lFolderID = GetNewFolderID(sPath)

                ' Add the file to the collection
                tPictureFiles.Add ReturnCPicture(sPath & sResult, bType, lFolderID)
            End If
        ElseIf bRecursive = True Then
            ' Search it, if it is a folder and the user
            ' does want to search recursive directories
            SearchForPics sPath & sResult & "\", bRecursive
        End If
    End If

    ' Find the next file
    lReturnVal = FindNextFile(hFindHandle, tFileData)
Loop

' Close file search
lReturnVal = FindClose(hFindHandle)

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function GetFolder() As String

' Purpose: Show the "Browser For Folder" window
'   and return the result

On Error GoTo ErrorHandler

Dim tFolder As BROWSEINFO
Dim sPath As String
Dim lFolderPIDL As Long
Dim iPosition As Integer

' Set the title text for the "Browser For Folder" window
With tFolder
    .lpszTitle = "Please select a folder that contains the pictures you would like to scroll through."
    .hwndOwner = Me.hwnd
End With

' Show the "Browser For Folder" window
lFolderPIDL = SHBrowseForFolder(tFolder)

' If the user didn't click cancel...
If lFolderPIDL <> 0 Then
    sPath = String(tProgramOptions.MAX_LEN, 0)

    ' Retrieve a "true" path from the PIDL number
    SHGetPathFromIDList lFolderPIDL, sPath

    iPosition = InStr(sPath, vbNullChar)

    ' Remove the last null character if there is one
    ' and set "sPath" to that value
    If iPosition <> 0 Then
        GetFolder = Left$(sPath, iPosition - 1)
    Else
        GetFolder = sPath
    End If
End If

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub AddFilesToList(ByVal sFolder As String)

' Purpose: Add the correct information to the listbox

On Error GoTo ErrorHandler

Dim nIndex As Long
Dim lFolderID As Long
Dim sPath As String
Dim sTemp As String
Dim iLastSlash As Integer

Static bAddingItems As Byte

If bAddingItems = True Then Exit Sub

Screen.MousePointer = 11

bAddingItems = True

' Set the folder that is currently shown ("" if
' all the folders are in view).
sCurrentFolder = sFolder

DoEvents

UpdateDirButtons

If sCurrentFolder <> "" And tPictureFiles.Count <> 0 Then
    txtLocation.Text = " (0)  " & sCurrentFolder

    lstTemp.ColumnHeaders(1).Text = "File Name"
    lstTemp.ColumnHeaders(1).Width = 3000
    lstTemp.ColumnHeaders(2).Text = "Size"
    lstTemp.ColumnHeaders(2).Width = 1000
Else
    txtLocation.Text = " (0)  All Folders"

    lstTemp.ColumnHeaders(1).Text = "Folder Name"
    lstTemp.ColumnHeaders(1).Width = 2000
    lstTemp.ColumnHeaders(2).Text = "Parent"
    lstTemp.ColumnHeaders(2).Width = 2000
End If

DoEvents

lstFiles.Visible = False

DoEvents

lstFiles.ListItems.Clear

DoEvents

If sCurrentFolder <> "" And tPictureFiles.Count <> 0 Then
    lstFiles.ColumnHeaders(1).Text = "File Name"
    lstFiles.ColumnHeaders(1).Width = 3000
    lstFiles.ColumnHeaders(2).Text = "Size"
    lstFiles.ColumnHeaders(2).Width = 1000

    ' Allow the user to go to the main list
    lstFiles.ListItems.Add , "UP_LEVEL", "..", , 2

    ' Get the folder ID of this folder
    lFolderID = GetFolderID(sCurrentFolder)

    For nIndex = 1 To tPictureFiles.Count
        If bUnloadNow = True Then Exit Sub

        If nIndex Mod 50 = 0 Then
            txtLocation.Text = " (" & lstFiles.ListItems.Count - 1 & ")  " & sCurrentFolder
        End If

        If tPictureFiles(nIndex).FolderID = lFolderID Then
            sPath = tPictureFiles(nIndex).FileName

            iLastSlash = LastSlash(sPath)

            lstFiles.ListItems.Add , "PIC_" & nIndex, Right$(sPath, Len(sPath) - iLastSlash), , tPictureFiles(nIndex).PicType + 2

            ' If the folder is a root drive, then
            ' don't remove the slash
            sTemp = Left$(sPath, iLastSlash)
            If Len(sTemp) <> 3 Then sTemp = Left$(sTemp, Len(sTemp) - 1)

            lstFiles.ListItems("PIC_" & nIndex).SubItems(1) = Round(FileLen(sPath) \ 1000) & "kb"
            lstFiles.ListItems("PIC_" & nIndex).bOld = True
        End If

        If bUnloadNow = True Then Exit Sub

        If bCancelOp = True Then
            GoTo DoneAdding
        Else: DoEvents
        End If
    Next nIndex
Else
    lstFiles.ColumnHeaders(1).Text = "Folder Name"
    lstFiles.ColumnHeaders(1).Width = 2000
    lstFiles.ColumnHeaders(2).Text = "Parent"
    lstFiles.ColumnHeaders(2).Width = 2000

    For nIndex = 1 To tAddedFolders.Count
        If bUnloadNow = True Then Exit Sub

        If nIndex Mod 50 = 0 Then
            txtLocation.Text = " (" & lstFiles.ListItems.Count & ")  All Folders"
        End If

        sPath = tAddedFolders(nIndex).FolderName

        ' Remove the last slash from the path only if
        ' it isn't a root drive
        If Len(sPath) <> 3 Then sPath = Left$(sPath, Len(sPath) - 1)

        ' If the path is a root drive, then
        ' only put the path in the left column; if
        ' it refers to a directory then put the
        ' folder name in the left side and it's
        ' parent's path in the second column
        If Len(sPath) <> 3 Then
            iLastSlash = LastSlash(sPath)

            lstFiles.ListItems.Add , "DIR_" & tAddedFolders(nIndex).FolderID, Right$(sPath, Len(sPath) - iLastSlash), , 1

            ' Get the parent path, but remove the last
            ' slash if it isn't a root drive
            sPath = Left$(sPath, iLastSlash)
            If Len(sPath) <> 3 Then sPath = Left$(sPath, Len(sPath) - 1)

            lstFiles.ListItems(lstFiles.ListItems.Count).SubItems(1) = sPath
        Else: lstFiles.ListItems.Add , "DIR_" & tAddedFolders(nIndex).FolderID, sPath, , 1
        End If

        lstFiles.ListItems(lstFiles.ListItems.Count).bOld = True

        If bUnloadNow = True Then Exit Sub

        If bCancelOp = True Then
            GoTo DoneAdding
        Else: DoEvents
        End If
    Next nIndex
End If

lstFiles.Visible = True

DoneAdding:

If lstFiles.ListItems.Count = 0 Then lstFiles.ListItems.Add , "NO_FILES", "No files.", , 2

UpdateDirButtons

' Put that in a format the user can understand.
If sCurrentFolder = "" Then
    txtLocation.Text = " (" & lstFiles.ListItems.Count & ")  All Folders"
Else: txtLocation.Text = " (" & lstFiles.ListItems.Count - 1 & ")  " & sCurrentFolder
End If

bAddingItems = False
bCancelOp = False

Screen.MousePointer = 0

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub GetPicIcons()

' Purpose: Load the system icons of different pictures

On Error GoTo ErrorHandler

Dim tPicIcon As SHFILEINFO

SHGetFileInfo ".bmp", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".jpg", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".gif", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".wmf", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".emf", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".ico", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".pcx", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".psd", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".tga", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

SHGetFileInfo ".lbm", vbNormal, tPicIcon, Len(tPicIcon), SHGFI_USEFILEATTRIBUTES Or SHGFI_ICON Or SHGFI_SMALLICON
GoSub AddToImageList

picHidden.Picture = LoadPicture()

Exit Sub

AddToImageList:

With picHidden
    .Picture = LoadPicture()
    .Width = 16
    .Height = 16

    DrawIconEx .hdc, 0, 0, tPicIcon.hIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon tPicIcon.hIcon
End With

imglstIcons.ListImages.Add , , picHidden.Image

Return

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub MovePicture(ByVal sDirection As String)

' Purpose: Flips two items in the collection of pictures.

On Error GoTo ErrorHandler

If sDirection = "MOVE_UP" Then
    ' We can't move the first picture item
    ' up.  (Remember the first item isn't a
    ' picture, it's the "up level" item)
    If lstFiles.SelectedItem.Index > 2 Then
        FlipItems lstFiles.SelectedItem.Index - 1, lstFiles.SelectedItem.Index, True
    End If
ElseIf sDirection = "MOVE_DOWN" Then
    ' We can't move the first picture item
    ' up.  (Remember the first item isn't a
    ' picture, it's the "up level" item)
    If lstFiles.SelectedItem.Index < lstFiles.ListItems.Count Then
        FlipItems lstFiles.SelectedItem.Index + 1, lstFiles.SelectedItem.Index, True
    End If
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub MoveFolder(ByVal sDirection As String)

' Purpose: Move a folder of pictures up or down through
'   the collection of pictures; it also changes the
'   collection of folders.

On Error GoTo ErrorHandler

Dim sPath As String
Dim lOtherIndex As Long
Dim lSelectedID As Long
Dim lOtherID As Long
Dim nIndex1 As Long
Dim nIndex2 As Long
Dim tTempList As New Collection
Dim bFoundSeries As Byte
Dim sFolderName As String

If sDirection = "MOVE_UP" Then
    If lstFiles.SelectedItem.Index = 1 Then
        Exit Sub
    Else: lOtherIndex = lstFiles.SelectedItem.Index - 1
    End If
ElseIf sDirection = "MOVE_DOWN" Then
    If lstFiles.SelectedItem.Index = lstFiles.ListItems.Count Then
        Exit Sub
    Else: lOtherIndex = lstFiles.SelectedItem.Index + 1
    End If
End If

' Get the FolderID of the item (either ABOVE or BELOW)
lOtherID = GetFolderID(GetItemPath(lOtherIndex))

' Get the FolderID of the selected path
lSelectedID = GetFolderID(GetItemPath(lstFiles.SelectedItem.Index))

' First, store all the picture files that are in the
' selected folder; then, remove them from the collection
For nIndex1 = 1 To tPictureFiles.Count
    If tPictureFiles(nIndex1).FolderID = lSelectedID Then
        tTempList.Add tPictureFiles(nIndex1)
        tPictureFiles.Remove nIndex1

        nIndex1 = nIndex1 - 1

        If nIndex1 + 1 > tPictureFiles.Count Then Exit For

        bFoundSeries = True
    ElseIf bFoundSeries = True Then
        Exit For
    End If
Next nIndex1

bFoundSeries = False

If sDirection = "MOVE_UP" Then
    ' Find the index number of the first picture that
    ' is in the folder above the selected one
    For nIndex2 = nIndex1 To 1 Step -1
        If tPictureFiles(nIndex2).FolderID = lOtherID Then
            If nIndex1 <> 1 Then nIndex1 = nIndex1 - 1
            bFoundSeries = True
        ElseIf bFoundSeries = True Then
            nIndex1 = nIndex1 + 1
            Exit For
        End If
    Next nIndex2
ElseIf sDirection = "MOVE_DOWN" Then
    ' Find the index number of the last picture that
    ' is in the folder below the selected one
    For nIndex2 = nIndex1 To tPictureFiles.Count
        If tPictureFiles(nIndex2).FolderID = lOtherID Then
            nIndex1 = nIndex1 + 1
            bFoundSeries = True
        ElseIf bFoundSeries = True Then
            Exit For
        End If
    Next nIndex2

    nIndex1 = nIndex1 - 1
End If

' Add all the pictures in the selected folder back
' into the picture collection
For nIndex2 = tTempList.Count To 1 Step -1
    If sDirection = "MOVE_UP" Then
        ' Add the pictures before the first picture
        ' in the folder above the one selected
        tPictureFiles.Add tTempList(nIndex2), , nIndex1
    ElseIf sDirection = "MOVE_DOWN" Then
        ' Add the pictures after the last picture in
        ' the folder below the one selected
        tPictureFiles.Add tTempList(nIndex2), , , nIndex1
    End If
Next nIndex2

' Store the path of the folder for each item involved
sFolderName = tAddedFolders("DIR_" & lSelectedID).FolderName

' First remove both of them from the collection
tAddedFolders.Remove "DIR_" & lSelectedID

If sDirection = "MOVE_UP" Then
    tAddedFolders.Add ReturnCFolder(sFolderName, lSelectedID), "DIR_" & lSelectedID, "DIR_" & lOtherID
ElseIf sDirection = "MOVE_DOWN" Then
    tAddedFolders.Add ReturnCFolder(sFolderName, lSelectedID), "DIR_" & lSelectedID, , "DIR_" & lOtherID
End If

' Invert the listbox items so the user knows
' what happened.
FlipItems lOtherIndex, lstFiles.SelectedItem.Index, False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub FlipItems(ByVal lFirstItem As Long, ByVal lSecondItem As Long, ByVal bDoCollection As Byte)

' Purpose: Flip two items in the lstFiles.  Optionally
'   will also flip the two items in the picture collection

On Error GoTo ErrorHandler

Dim sTempText As String
Dim sTempSubItem As String
Dim bTempIcon As Byte
Dim sTempKey1 As String
Dim sTempKey2 As String

' Store the info of the first item.
With lstFiles.ListItems(lFirstItem)
    sTempText = .Text
    sTempSubItem = .SubItems(1)
    bTempIcon = .SmallIcon
End With

' Set the first item's info to that of the second item.
With lstFiles.ListItems(lFirstItem)
    .Text = lstFiles.ListItems(lSecondItem).Text
    .SubItems(1) = lstFiles.ListItems(lSecondItem).SubItems(1)
    .SmallIcon = lstFiles.ListItems(lSecondItem).SmallIcon
End With

' Set the second item's info to that of the first.
With lstFiles.ListItems(lSecondItem)
    .Text = sTempText
    .SubItems(1) = sTempSubItem
    .SmallIcon = bTempIcon
End With

' Move the selection with the item
lstFiles.ListItems(lFirstItem).Selected = True
lstFiles.ListItems(lSecondItem).Selected = False

If bDoCollection = True Then
    ' Get the keys of both items.
    sTempKey1 = lstFiles.ListItems(lFirstItem).Key
    sTempKey2 = lstFiles.ListItems(lSecondItem).Key

    sTempKey1 = Right$(sTempKey1, Len(sTempKey1) - 4)
    sTempKey2 = Right$(sTempKey2, Len(sTempKey2) - 4)

    ' Store the information of the first item.
    With tPictureFiles(CLng(sTempKey1))
        sTempText = .FileName
        sTempSubItem = .PicType

        ' Change the information with that of the second item.
        .FileName = tPictureFiles(CLng(sTempKey2)).FileName
        .PicType = tPictureFiles(CLng(sTempKey2)).PicType
    End With

    ' Change the information in the second item with that
    ' that was in the first item.
    With tPictureFiles(CLng(sTempKey2))
        .FileName = sTempText
        .PicType = sTempSubItem
    End With
Else
    ' Store the two keys.
    sTempKey1 = lstFiles.ListItems(lFirstItem).Key
    sTempKey2 = lstFiles.ListItems(lSecondItem).Key

    ' Clear the two keys.
    lstFiles.ListItems(lFirstItem).Key = ""
    lstFiles.ListItems(lSecondItem).Key = ""

    ' Return the keys in reverse order.
    lstFiles.ListItems(lFirstItem).Key = sTempKey2
    lstFiles.ListItems(lSecondItem).Key = sTempKey1
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function GetItemPath(ByVal lItemIndex As Long) As String

' Purpose: Figure out the path of an item in the listbox
'   This is only meant to be used on directory items

On Error GoTo ErrorHandler

Dim sItemKey As String
Dim lItemID As Long
Dim sType As String

' Get the item's key.
sItemKey = lstFiles.ListItems(lItemIndex).Key
' Extract the ID of the item, either to a file or a folder.
lItemID = Right$(sItemKey, Len(sItemKey) - 4)
' Find out what the ID refers to.
sType = Left$(sItemKey, 4)

' Retrieve the file name of the specified item.
Select Case sType
    Case "PIC_"
        GetItemPath = tPictureFiles(lItemID).FileName
    Case "DIR_"
        GetItemPath = tAddedFolders("DIR_" & lItemID).FolderName
End Select

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub SetMusicType(ByVal bType As Byte)

' Purpose: Setup the music type the user selected

On Error GoTo ErrorHandler

Dim nIndex As Byte

' Find the checkbox that is currently checked and set
' it's tag property to identify it as the last clicked
For nIndex = 0 To 2
    If chkMusicType(nIndex).Value = 1 Then
        chkMusicType(nIndex).Tag = "LAST"
        Exit For
    Else: chkMusicType(nIndex).Tag = ""
    End If
Next nIndex

' Setup options necessary for the type selected
With tProgramOptions
    Select Case bType
        Case 1
            ' Fill in the text box with the music filename
            txtMusicFile.Text = .sMusicFile

            optLoopMusic(.bLoopMusic).Value = True
        Case 2
            optPlayTrack(.bPlayTrack).Value = True

            ' Fill the listbox with all the tracks
            FillTracks
    End Select
End With

' Set the value of each of the music type check boxes
' as well as show and hide the correct picture boxes
bByCode = True

For nIndex = 0 To 2
    chkMusicType(nIndex).Value = IIf(bType = nIndex, 1, 0)
    picMusicType(nIndex).Visible = IIf(bType = nIndex, True, False)
Next nIndex

bByCode = False

tProgramOptions.bMusicType = bType

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub FillTracks()

' Purpose: See if a CD is in the drive and if so,
'   add the tracks to cmbTracks

On Error GoTo ErrorHandler

Dim nIndex As Integer

' See if there's a CD is the drive
If BackMusic.GetTrackCount = True Then
    ' Clear the combo box of all previous tracks
    cmbTracks.Clear

    ' Add the tracks to the combo box
    For nIndex = 1 To BackMusic.NumOfTracks
        cmbTracks.AddItem "Track " & nIndex
    Next nIndex

    ' See if the user has previously set to
    ' play a specific track, and that that
    ' track number is not invalid for this CD.
    If tProgramOptions.iTrackNumber <> 0 And tProgramOptions.iTrackNumber <= cmbTracks.ListCount Then
        cmbTracks.ListIndex = tProgramOptions.iTrackNumber - 1
    Else: cmbTracks.ListIndex = 0
    End If

    ' Make sure the "Error" frame is hidden
    fraError.Visible = False
    fraNormal.Visible = True
Else
    fraError.Visible = True
    fraNormal.Visible = False
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub SetUserOptions()

' Purpose: Sets the value of most of the controls
'   to the user's last used value.

On Error GoTo ErrorHandler

' Interval Tab
Dim iPosX As Integer
Dim iPosY As Integer
Dim iRadius As Integer
Dim dRadians As Double
Dim dActual As Double

Dim bNoTransitions As Byte
Dim nIndex As Integer
Dim sTransition As String
Dim nRemove As Integer
Dim bModeFound As Byte

With tProgramOptions
    ' If this is the first start, then set the defaults.
    If tProgramOptions.bNotFirstStart = False Then SetDefaults

    ' ---------------------------------------------------
    ' Interval Tab

    iRadius = (picDial.ScaleWidth \ 2) - 10

    nIndex = 0

    For dRadians = 7.85398163397448 To 1.88495559215388 Step -0.314159265358
        If dRadians >= 6.28318530717959 Then
            dActual = dRadians - 6.28318530717959
        Else: dActual = dRadians
        End If

        iPosX = (iRadius * Cos(dActual)) + (0 * Sin(dActual))
        iPosY = (0 * Cos(dActual)) - (iRadius * Sin(dActual))

        iCirclePoints(nIndex).X = iPosX + iRadius + 10
        iCirclePoints(nIndex).Y = iPosY + iRadius + 10

        picDial.PSet (iPosX + iRadius + 10, iPosY + iRadius + 10)

        nIndex = nIndex + 1
    Next

    iPoint = (.iInterval \ 1000) - 1

    picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    txtTime.Text = .iInterval \ 1000

    ' ---------------------------------------------------
    ' Transitions Tab

    ' See if the user has previously set the
    ' transitions to use.
    On Error Resume Next
    If UBound(.bTransitions) = 0 Then
        ' If an error occurs this will also be set.
        bNoTransitions = True
    End If
    On Error GoTo ErrorHandler

    For nIndex = 1 To NUM_OF_TRANSITIONS
        lstAvailable.AddItem LookUpTrans(nIndex)
        lstAvailable.ItemData(nIndex - 1) = nIndex
    Next nIndex

    If bNoTransitions = False Then
        For nIndex = 1 To UBound(.bTransitions)
            sTransition = LookUpTrans(.bTransitions(nIndex))

            lstActual.AddItem sTransition
            lstActual.ItemData(lstActual.ListCount - 1) = .bTransitions(nIndex)

            For nRemove = 0 To lstAvailable.ListCount
                If lstAvailable.List(nRemove) = sTransition Then
                    lstAvailable.RemoveItem nRemove
                    Exit For
                End If
            Next nRemove
        Next nIndex
    End If

    If lstAvailable.ListCount <> 0 Then lstAvailable.ListIndex = 0
    If lstActual.ListCount <> 0 Then lstActual.ListIndex = 0

    chkRandom.Value = .bRandomTransitions

    UpdateTransButtons

    ' ---------------------------------------------------
    ' Miscellaneous Tab

    ' Scroll Direction
    optDirection(.bScrollDirection).Value = True
    ' Picture Size
    cmbPictureSize.ListIndex = .bPictureSize

    ' Sound Effects
    chkSoundEffects.Value = .bSoundEffects

    ' ---------------------------------------------------
    ' Advanced Tab

    ' Display Mode Settings
    optScreenSize(.bScreenSetting).Value = True

    ' Retrieve all supported display modes
    DirectDraw.GetDisplayModes Me.hwnd

    cmbDisplayModes.Clear

    ' Add all supported options to the list.
    For nIndex = 1 To DirectDraw.ModeCount
        cmbDisplayModes.AddItem DirectDraw.ModeWidth(nIndex) & " x " & DirectDraw.ModeHeight(nIndex) & " x " & DirectDraw.ModeBPP(nIndex)
    Next nIndex

    ' If the user has previously set a display
    ' mode to use, then find it and set the
    ' selected item to be it.
    With .tDisplayMode
        If .iWidth <> 0 And .iHeight <> 0 And .bBPP <> 0 Then
            For nIndex = 1 To DirectDraw.ModeCount
                If .iWidth = DirectDraw.ModeWidth(nIndex) And .iHeight = DirectDraw.ModeHeight(nIndex) And .bBPP = DirectDraw.ModeBPP(nIndex) Then
                    cmbDisplayModes.ListIndex = nIndex - 1

                    bModeFound = True
                    Exit For
                End If
            Next nIndex
        End If
    End With

    ' If the user hasn't previously set a display
    ' mode, then just select the first one.
    If bModeFound = False Then cmbDisplayModes.ListIndex = 0

    ' Scrolling Colors
    picBackColor.BackColor = .lBackColor
    picInfoColor.BackColor = .lInfoColor
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub DrawTip(ByVal lTipID As Long, Optional ByVal sString As String)

' Purpose: Receives a string of text to print
'   into picInfo in the exact position.

On Error GoTo ErrorHandler

Dim sTip As String
Dim rInfo As RECT
Dim fHeight As Single

Const MARGIN = 5

If sString = "" Then
    sTip = LoadResString(lTipID)
Else: sTip = sString
End If

If picInfo.TextWidth(sTip) > (picInfo.Width - MARGIN) Then
    fHeight = picInfo.TextWidth(sTip) / (picInfo.Width - MARGIN)

    If fHeight > 1 And fHeight < 2 Then fHeight = Fix(fHeight + 1)

    fHeight = fHeight * picInfo.TextHeight(sTip)
Else: fHeight = picInfo.TextHeight(sTip)
End If

' Prepare the dimensions for drawing tips
' in picInfo
With rInfo
    .Left = MARGIN
    .Top = ((picInfo.Height - fHeight) \ 2) - 1
    .Right = picInfo.Width - MARGIN
    .Bottom = .Top + fHeight
End With

picInfo.Cls

' Draw the tip for the button
DrawTextAPI picInfo.hdc, sTip, Len(sTip), rInfo, DT_CENTER Or DT_WORDBREAK Or DT_NOCLIP

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function LookUpTrans(ByVal bTransition As Byte) As String

' Purpose: Give it a transition number and it
'   returns its description text.

On Error GoTo ErrorHandler

Select Case bTransition
    Case 1: LookUpTrans = "Arrange Pieces"
    Case 2: LookUpTrans = "Blinds Horizontal"
    Case 3: LookUpTrans = "Blinds Vertical"
    Case 4: LookUpTrans = "Box In"
    Case 5: LookUpTrans = "Box Out"
    Case 6: LookUpTrans = "Circle In"
    Case 7: LookUpTrans = "Circle Out"
    Case 8: LookUpTrans = "Cross In"
    Case 9: LookUpTrans = "Diagonal Slide Left Down"
    Case 10: LookUpTrans = "Diagonal Slide Right Down"
    Case 11: LookUpTrans = "Diagonal Slide Left Up"
    Case 12: LookUpTrans = "Diagonal Slide Right Up"
    Case 13: LookUpTrans = "Diagonal Squeeze Left"
    Case 14: LookUpTrans = "Diagonal Squeeze Right"
    Case 15: LookUpTrans = "Diagonal Squeeze Up #1"
    Case 16: LookUpTrans = "Diagonal Squeeze Up #2"
    Case 17: LookUpTrans = "Falling Lines Down"
    Case 18: LookUpTrans = "Falling Lines Left"
    Case 19: LookUpTrans = "Falling Lines Right"
    Case 20: LookUpTrans = "Falling Lines Up"
    Case 21: LookUpTrans = "Maze"
    Case 22: LookUpTrans = "Move Down"
    Case 23: LookUpTrans = "Move Left"
    Case 24: LookUpTrans = "Move Right"
    Case 25: LookUpTrans = "Move Up"
    Case 26: LookUpTrans = "Move Out Horizontal"
    Case 27: LookUpTrans = "Move Out Vertical"
    Case 28: LookUpTrans = "Puzzle"
    Case 29: LookUpTrans = "Random Dots (Small)"
    Case 30: LookUpTrans = "Random Dots (Medium)"
    Case 31: LookUpTrans = "Random Dots (Large)"
    Case 32: LookUpTrans = "Random Lines Horizontal"
    Case 33: LookUpTrans = "Random Lines Vertical"
    Case 34: LookUpTrans = "Sandwich Horizontal"
    Case 35: LookUpTrans = "Sandwich Vertical"
    Case 36: LookUpTrans = "Shades Down and Up"
    Case 37: LookUpTrans = "Shades Down Twice"
    Case 38: LookUpTrans = "Shades Right and Left"
    Case 39: LookUpTrans = "Shades Right Twice"
    Case 40: LookUpTrans = "Slide In"
    Case 41: LookUpTrans = "Smear"
    Case 42: LookUpTrans = "Snake Horizontal"
    Case 43: LookUpTrans = "Snake Vertical"
    Case 44: LookUpTrans = "Stretch Out Horizontal"
    Case 45: LookUpTrans = "Stretch Out Vertical"
    Case 46: LookUpTrans = "Tile Across"
    Case 47: LookUpTrans = "Wipe Down"
    Case 48: LookUpTrans = "Wipe Left"
    Case 49: LookUpTrans = "Wipe Right"
    Case 50: LookUpTrans = "Wipe Up"
End Select

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub tlbDirectory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ErrorHandler

Dim nIndex As Integer

For nIndex = 1 To tlbDirectory.buttons.Count
    If X >= tlbDirectory.buttons(nIndex).Left _
        And X <= tlbDirectory.buttons(nIndex).Left + tlbDirectory.ButtonWidth _
        And Y >= tlbDirectory.buttons(nIndex).Top _
        And Y <= tlbDirectory.buttons(nIndex).Top + tlbDirectory.ButtonHeight _
        And nIndex <> 3 And nIndex <> 6 Then
            DrawTip 119 + nIndex
            Exit For
    End If
Next nIndex

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tlbMoveDown_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

Dim sName As String
Dim sData As String

' Disable the DOWN button if moving this one down
' will make it the last item.
If lstActual.List(lstActual.ListIndex + 2) = "" Then tlbMoveDown.buttons(1).Enabled = False

' Enable the UP button if it is currently disabled.
If tlbMoveUp.buttons(1).Enabled = False Then tlbMoveUp.buttons(1).Enabled = True

' Store the info of the item before the selected.
sName = lstActual.List(lstActual.ListIndex + 1)
sData = lstActual.ItemData(lstActual.ListIndex + 1)

' Set the info the item before the selected to
' the info of this one.
lstActual.List(lstActual.ListIndex + 1) = lstActual.List(lstActual.ListIndex)
lstActual.ItemData(lstActual.ListIndex + 1) = lstActual.ItemData(lstActual.ListIndex)

' Set the info of the selected item to that of
' the item before this one.
lstActual.List(lstActual.ListIndex) = sName
lstActual.ItemData(lstActual.ListIndex) = sData

lstActual.ListIndex = lstActual.ListIndex + 1

UpdateTransitions

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tlbMoveDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 347

End Sub
Private Sub tlbMoveUp_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

Dim sName As String
Dim sData As String

' Disable the UP button if moving this one up
' will make it the first item.
If lstActual.List(lstActual.ListIndex - 2) = "" Then tlbMoveUp.buttons(1).Enabled = False

' Enable the DOWN button if it is currently disabled.
If tlbMoveDown.buttons(1).Enabled <> True Then tlbMoveDown.buttons(1).Enabled = True

' Store the info of the item before the selected.
sName = lstActual.List(lstActual.ListIndex - 1)
sData = lstActual.ItemData(lstActual.ListIndex - 1)

' Set the info the item before the selected to
' the info of this one.
lstActual.List(lstActual.ListIndex - 1) = lstActual.List(lstActual.ListIndex)
lstActual.ItemData(lstActual.ListIndex - 1) = lstActual.ItemData(lstActual.ListIndex)

' Set the info of the selected item to that of
' the item before this one.
lstActual.List(lstActual.ListIndex) = sName
lstActual.ItemData(lstActual.ListIndex) = sData

lstActual.ListIndex = lstActual.ListIndex - 1

UpdateTransitions

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tlbMoveUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 346

End Sub
Private Sub UpdateTransButtons()

' Purpose: On the transitions tab, updates the
'   various buttons to reflex the status of the
'   listboxes and the user's selection.

On Error GoTo ErrorHandler

If lstActual.ListIndex > 0 And tProgramOptions.bRandomTransitions = False Then
    tlbMoveUp.buttons(1).Enabled = True
Else: tlbMoveUp.buttons(1).Enabled = False
End If

If lstActual.ListIndex < lstActual.ListCount - 1 And lstActual.ListIndex <> -1 And tProgramOptions.bRandomTransitions = False Then
    tlbMoveDown.buttons(1).Enabled = True
Else: tlbMoveDown.buttons(1).Enabled = False
End If

If lstActual.ListCount > 0 Then
    cmdRemoveAll.Enabled = True
Else: cmdRemoveAll.Enabled = False
End If

If lstAvailable.ListCount = 0 Then
    cmdAddAll.Enabled = False
Else: cmdAddAll.Enabled = True
End If

If lstActual.ListIndex <> -1 Then
    cmdRemove.Enabled = True
Else: cmdRemove.Enabled = False
End If

If lstAvailable.ListIndex <> -1 Then
    cmdAdd.Enabled = True
Else: cmdAdd.Enabled = False
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub UpdateTransitions()

' Purpose: Recreates the user's settings for the
'   transitions to use.

On Error GoTo ErrorHandler

Dim nIndex As Integer

With tProgramOptions
    If lstActual.ListCount = 0 Then
        ReDim .bTransitions(0)
    Else
        ReDim .bTransitions(1 To lstActual.ListCount)

        For nIndex = 1 To lstActual.ListCount
            .bTransitions(nIndex) = lstActual.ItemData(nIndex - 1)
        Next nIndex
    End If
End With

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub UpdateDirButtons()

' Purpose: On the "Select Directory" option, updates the
'   various buttons to reflex the status of the listbox
'   and the user's selection.

On Error GoTo ErrorHandler

' If there aren't any items just disable both buttons.
If lstFiles.ListItems.Count = 0 Then GoTo DisableBoth

If lstFiles.ListItems(1).Key = "NO_FILES" Then
    tlbDirectory.buttons("REMOVE").Enabled = False
Else: tlbDirectory.buttons("REMOVE").Enabled = True
End If

' If a file is selected but not more than one.
If Not lstFiles.SelectedItem Is Nothing Then
    ' Depending on the view (folders or files), we
    ' can allow the item selected to be moved.
    ' If folders, then from the first item in the list can
    ' they be moved down; otherwise, only from the second
    ' item can they be moved.
    If lstFiles.SelectedItem.Index > IIf(sCurrentFolder = "", 0, 1) And lstFiles.SelectedItem.Index < lstFiles.ListItems.Count Then
        tlbDirectory.buttons(4).Enabled = True
    Else: tlbDirectory.buttons(4).Enabled = False
    End If

    ' Depending on the view (folders or files), we
    ' can allow the item selected to be moved.
    ' If folders, then from the second item in the list
    ' can they be moved up; otherwise, only from the
        ' third item they be moved.
    If lstFiles.SelectedItem.Index > IIf(sCurrentFolder = "", 1, 2) Then
        tlbDirectory.buttons(5).Enabled = True
    Else: tlbDirectory.buttons(5).Enabled = False
    End If
Else: GoTo DisableBoth
End If

Exit Sub

DisableBoth:
' Disable the buttons if more than one item is
' selected, if no item is selected, or if there aren't
' any items in the list box.
tlbDirectory.buttons(4).Enabled = False
tlbDirectory.buttons(5).Enabled = False

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub RemoveItems()

' Purpose: Remove either the selected pictures from
'   the collection or a whole folder of pictures.

On Error GoTo ErrorHandler

Dim nIndexDirs As Long
Dim nIndexFiles As Long
Dim lFolderID As Long
Dim bUp As Byte

If sCurrentFolder = "" Then
    If MsgBox("Are you sure you want to remove ALL the pictures that are in the directories that you have selected?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    ' Show the wait window and start removing pictures
    WaitProcess "Removing Pictures", False, Me

    For nIndexDirs = 1 To lstFiles.ListItems.Count
        If nIndexDirs > lstFiles.ListItems.Count Then Exit For

        If lstFiles.ListItems(nIndexDirs).Selected = True Then
            ' Loop through the collection of picture files
            ' and remove all those from the selected folder
            nIndexFiles = 1

            lFolderID = GetFolderID(GetItemPath(nIndexDirs))

            Do While nIndexFiles <= tPictureFiles.Count
                ' If this file is in this directory, remove
                ' it but don't increase the index variable:
                ' we just removed something so the index
                ' already dropped itself once.
                If tPictureFiles(nIndexFiles).FolderID = lFolderID Then
                    tPictureFiles.Remove nIndexFiles
                Else: nIndexFiles = nIndexFiles + 1
                End If

                DoEvents
            Loop

            ' Remove the folder from the folder collection
            tAddedFolders.Remove "DIR_" & lFolderID

            ' Remove the folder from the list
            lstFiles.ListItems.Remove "DIR_" & lFolderID

            nIndexDirs = nIndexDirs - 1
        End If
    Next nIndexDirs

    AddFilesToList ""

    EndWaitProcess Me
Else
    For nIndexFiles = lstFiles.ListItems.Count To 1 Step -1
        If lstFiles.ListItems(nIndexFiles).Selected = True Then
            nIndexDirs = nIndexDirs + 1

            If lstFiles.ListItems(nIndexFiles).Key = "UP_LEVEL" Then bUp = 1
        End If
    Next nIndexFiles

    If bUp = 1 And nIndexDirs = 1 Then Exit Sub

    If MsgBox("Are you sure you want to remove the selected file(s) from the list of pictures?", vbYesNo + vbExclamation) = vbNo Then Exit Sub

    Screen.MousePointer = 11

    For nIndexFiles = lstFiles.ListItems.Count To 1 Step -1
        If lstFiles.ListItems(nIndexFiles).Selected = True Then
            ' If this item is the "up level" item at the top
            ' remove it from the list but not from the collection.
            If lstFiles.ListItems(nIndexFiles).Key <> "UP_LEVEL" Then
                ' Remove the picture from the collection.
                ' We must use the key property which
                ' corresponds with the collection, but we
                ' need to remove the # at the front
                tPictureFiles.Remove CLng(Right$(lstFiles.ListItems(nIndexFiles).Key, Len(lstFiles.ListItems(nIndexFiles).Key) - 4))
            End If

            lstFiles.ListItems.Remove nIndexFiles
        End If
    Next nIndexFiles

    picPreview.Picture = LoadPicture()

    ' Remove the folder as well if there aren't
    ' any files left
    If lstFiles.ListItems.Count = 1 Then
        tAddedFolders.Remove "DIR_" & GetFolderID(sCurrentFolder)

        AddFilesToList ""
    Else: AddFilesToList sCurrentFolder
    End If

    Screen.MousePointer = 0
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrFlash_Timer()

On Error GoTo ErrorHandler

If bFlashCount < 7 Then
    If bFlashCount Mod 2 <> 0 Then
        lblNoPics.ForeColor = vbRed
    Else: lblNoPics.ForeColor = vbButtonText
    End If

    bFlashCount = bFlashCount + 1
Else
    tmrFlash.Enabled = False

    lblNoPics.Visible = False
    bFlashCount = 0
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub tmrLoadPreview_Timer()

On Error GoTo ErrorHandler

Dim lPictureIndex As Long
Dim rPreview As RECT

tmrLoadPreview.Enabled = False

Screen.MousePointer = 11

' Extract the picture index from the key of the item.
lPictureIndex = Right$(lstFiles.SelectedItem.Key, Len(lstFiles.SelectedItem.Key) - 4)

' Attempt to load the picture.
If GetPicFromIndex(picHidden, lPictureIndex, picPreview.Width, picPreview.Height, True) = False Then
    With rPreview
        .Top = (106 - (picPreview.TextHeight("A") * 2 + 10)) / 2
        .Right = 106
        .Bottom = .Top + picPreview.TextHeight("A")
    End With

    picHidden.Picture = LoadPicture()
    DrawTextAPI picHidden.hdc, "Cannot", 6, rPreview, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

    rPreview.Top = rPreview.Top + picPreview.TextHeight("A") + 10
    rPreview.Bottom = rPreview.Top + picPreview.TextHeight("A")

    DrawTextAPI picHidden.hdc, "Load Picture", 12, rPreview, DT_CENTER Or DT_NOCLIP
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
Private Sub txtMusicFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 373

End Sub
Private Sub txtMusicFile_Validate(Cancel As Boolean)

On Error GoTo ErrorHandler

If Dir(txtMusicFile.Text) = "" Then
    If MsgBox("Please enter a valid path to a music file.", vbOKCancel + vbExclamation) = vbCancel Then
        txtMusicFile.Text = tProgramOptions.sMusicFile
    End If

    Cancel = True
Else: tProgramOptions.sMusicFile = txtMusicFile.Text
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub MenuSelect(ByVal hwnd As Long)

' Purpose: This gets called whenever the user's mouse
'   goes over a menu item.

On Error GoTo ErrorHandler

Dim hMenu As Long
Dim hSubMenu As Long
Dim iMenuCount As Integer
Dim nIndex As Integer

' Get the menu ID
hMenu = GetMenu(hwnd)

If hMenu <> 0 Then
    ' Get the number of main menu items.
    iMenuCount = GetMenuItemCount(hMenu)

    For nIndex = 0 To iMenuCount - 1
        ' If this menu item is selected, then walk
        ' through it to find the one actually selected.
        If GetMenuState(hMenu, nIndex, MF_BYPOSITION And MF_HILITE) Then
            hSubMenu = GetSubMenu(hMenu, nIndex)
            WalkSubMenu hSubMenu
        End If
    Next nIndex
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub WalkSubMenu(hSubMenu As Long)

' Purpose: Used to "walk" through a sub menu to
'   find the one selected.

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim iMenuCount As Integer
Dim hSubSubMenu As Long
Dim sMenuCaption As String
Dim lReturnVal As Long

' Get the count of menu items in this menu.
iMenuCount = GetMenuItemCount(hSubMenu)

' Loop through all the items on the menu.
For nIndex = 0 To iMenuCount - 1
    ' Determine whether this item is highlighted.
    If GetMenuState(hSubMenu, nIndex, MF_BYPOSITION) And MF_HILITE Then
        ' Attempt to get a submenu of this item.
        hSubSubMenu = GetSubMenu(hSubMenu, nIndex)

        ' Check for a submenu with an item selected.
        If hSubSubMenu <> 0 And AnyLit(hSubSubMenu) = True Then
            ' There is a submenu with a selection so walk it.
            WalkSubMenu hSubSubMenu

            Exit For
        Else
            ' This is it: the user's selection.

            ' Set buffer size.
            sMenuCaption = Space(255)

            ' Retrieve the menu's caption.
            lReturnVal = GetMenuString(hSubMenu, nIndex, sMenuCaption, Len(sMenuCaption), MF_BYPOSITION)

            ' Trim the buffer of extra characters.
            sMenuCaption = Left$(sMenuCaption, lReturnVal)

            Select Case sMenuCaption
                Case "&Directory"
                    DrawTip 50
                Case "&File(s)"
                    DrawTip 51
                Case "&Load Saved List"
                    DrawTip 52
                Case "&Save Current List"
                    DrawTip 53
                Case Else
                    DrawTip 0, frmMenu.mnuAddItem(nIndex).Tag
            End Select

            Exit Sub
        End If
    End If
Next nIndex

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function AnyLit(hSubMenu As Long) As Byte

' Purpose: Search a menu to see if any items are selected.

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim iMenuCount As Long

' Get the number of items in the menu.
iMenuCount = GetMenuItemCount(hSubMenu)

' Loop through the menu items.
For nIndex = 0 To iMenuCount - 1
    ' Check whether this item is highlighted.
    If GetMenuState(hSubMenu, nIndex, MF_BYPOSITION) And MF_HILITE Then
        AnyLit = True
        Exit Function
    End If
Next nIndex

' Return FALSE, no items highlighted.
AnyLit = False

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub ShowPopMenu(ByVal objMenu As Menu, ByVal X As Integer, ByVal Y As Integer)

' Purpose: Show a popup menu.  Since the menu will be
'   from frmMenu, unload frmMenu when the popup menu
'   is done.  This will stop subclassing.

On Error GoTo ErrorHandler

Me.PopupMenu objMenu, , X, Y
Unload frmMenu

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Public Sub AddMenuClick(ByVal iItemIndex As Byte)

' Purpose: Handles when the user clicks an item in
'   the "Add" menu.

On Error GoTo ErrorHandler

Dim sFolder As String
Dim sSelectedFiles() As String
Dim lFolderID As Long
Dim bRecursive As Byte
Dim bType As Byte
Dim nIndex As Integer

Select Case iItemIndex
    ' If the user clicked add directory...
    Case 0
        sFolder = GetFolder

        If sFolder = "" Then Exit Sub

        NormalizePath sFolder

        bRecursive = MsgBox("Would you like to include the pictures that may be in any sub folders as well?", vbYesNoCancel + vbQuestion)

        Select Case bRecursive
            Case vbYes
                bRecursive = True
            Case vbNo
                bRecursive = False
            Case vbCancel
                Exit Sub
        End Select

        ' Show the Wait window
        WaitProcess "Searching for pictures", True, Me

        ' Search for pictures.
        SearchForPics sFolder, bRecursive

        ' Stop the wait process.
        EndWaitProcess Me

        AddFilesToList ""

    ' If the user clicked add files...
    Case 1
        sFolder = OpenFileDialog(Me.hwnd, PICTURES, sSelectedFiles)

        If sSelectedFiles(0) = "" Then Exit Sub

        Screen.MousePointer = 11

        ' Get a unique folder ID.  If the user has previously
        ' added pictures from this directory, that's what
        ' we want.
        lFolderID = GetNewFolderID(sFolder)

        ' Add the files to the collection of picture files
        For nIndex = 0 To UBound(sSelectedFiles)
            bType = ConfirmType(sSelectedFiles(nIndex))

            If bType <> 0 Then tPictureFiles.Add ReturnCPicture(sSelectedFiles(nIndex), bType, lFolderID)
        Next nIndex

        AddFilesToList ""

        Screen.MousePointer = 0

    ' If the user clicked load...
    Case 3
GetFileName:
        ' Ask the user for a file.
        OpenFileDialog Me.hwnd, SAVED_LIST, sSelectedFiles

        ' If the user didn't press cancel.
        If sSelectedFiles(0) <> "" Then
            WaitProcess "Loading List", True, Me

            ' Attempt to load the saved list.
            If LoadSavedList(sSelectedFiles(0), True, True) = False Then
                GoSub LoadErrMsg
                GoTo GetFileName
            End If

            AddFilesToList ""
        End If

        EndWaitProcess Me

    ' If the user clicked save...
    Case 4
        sFolder = SaveFileDialog(Me.hwnd)

        If sFolder <> "" Then
            SaveCurrentList sFolder
        End If
    Case Is > 5
        WaitProcess "Loading List", True, Me

        ' Attempt to load the saved list.
        If LoadSavedList(IIf(iItemIndex = 6, tProgramOptions.sRecent1, tProgramOptions.sRecent2), True, True) = False Then
            GoSub LoadErrMsg
            GoTo GetFileName
        End If

        EndWaitProcess Me

        AddFilesToList ""
End Select

Exit Sub

LoadErrMsg:
' Display a message if we cannot open the list.
MsgBox "This file is:" & vbCr & "1. Not available now, or" & vbCr & "2. Not in the correct format for loading into Picture Scroller" & vbCr & vbCr & "Please select a different file.", vbInformation
Return

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub SaveCurrentList(ByVal sFileName As String)

On Error GoTo ErrorHandler

Dim nIndex As Long

Open sFileName For Output As FILENUM_LIST

Print #FILENUM_LIST, LIST_HEADER

For nIndex = 1 To tPictureFiles.Count
    Print #FILENUM_LIST, tPictureFiles(nIndex).FileName
Next nIndex

Close FILENUM_LIST

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub UpdatePreview()

On Error GoTo ErrorHandler

BitBlt picShow.hdc, (picShow.ScaleWidth - picSlider.ScaleWidth) / 2, 0, picShow.ScaleWidth, picShow.ScaleHeight, picSlider.hdc, 0, iYPos, vbSrcCopy

picShow.Refresh

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub PrintOnDC(ByVal bSlot As Byte)

On Error GoTo ErrorHandler

Dim iPosX As Integer

iPosX = 10

If bSlot Mod 2 = 0 Then iPosX = iPosX + iSize + 10

BitBlt picSlider.hdc, iPosX + (iSize - picHidden.Width) / 2, (((bSlot - 1) \ 2) * (iSize + 10)) + 10 + (iSize - picHidden.Height) / 2, iSize, iSize, picHidden.hdc, 0, 0, vbSrcCopy

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function HitEdge() As Byte

On Error GoTo ErrorHandler

Dim bRow As Byte

bRow = bNextFree \ 2

If iYPos + picShow.ScaleHeight >= (iSize * bRow) + (bRow * 10) + 10 Then HitEdge = True

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub LoadRow(ByVal bDirection As Byte, ByVal bUpdate As Byte)

On Error GoTo ErrorHandler

Dim nIndex As Byte

If bNextFree = 0 Then bNextFree = 1

For nIndex = 1 To 2
    If bNextFree > 1 Then
        If lPreviewPics(bNextFree - 1) = tPictureFiles.Count Then Exit For
    End If

    If bDirection = DIR_UP Then
        lPreviewPics(nIndex) = lPreviewPics(nIndex) - 1
    ElseIf bDirection = DIR_DOWN Then
        If bNextFree <> 1 Then
            lPreviewPics(bNextFree) = lPreviewPics(bNextFree - 1) + 1
        Else: lPreviewPics(bNextFree) = lPreviewPics(bNextFree) + 1
        End If
    End If

    If lPreviewPics(bNextFree) - 1 < tPictureFiles.Count Then
        GetPicFromIndex picHidden, lPreviewPics(bNextFree), iSize, iSize, True

        PrintOnDC bNextFree

        bNextFree = bNextFree + 1

        If bUpdate = 1 Then UpdatePreview
    End If
Next nIndex

picHidden.Picture = LoadPicture()

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub AdjustRows(ByVal bDirection As Byte)

On Error GoTo ErrorHandler

Dim nIndex As Integer

If bDirection = DIR_UP And lPreviewPics(NUM_OF_PREVIEWS) <> tPictureFiles.Count Then
    BitBlt picSlider.hdc, 10, 10, (iSize * 2) + 10, picSlider.ScaleHeight - iSize - 10, picSlider.hdc, 10, iSize + 20, vbSrcCopy
    iYPos = ((iSize * ((NUM_OF_PREVIEWS \ 2) - 1)) + (NUM_OF_PREVIEWS \ 2) * 10) - picShow.ScaleHeight

    picSlider.Line (0, picSlider.Height - iSize - 10)-(picSlider.Width, picSlider.Height), vbWindowBackground, BF

    For nIndex = 1 To NUM_OF_PREVIEWS - 2
        lPreviewPics(nIndex) = lPreviewPics(nIndex + 2)
    Next nIndex
ElseIf bDirection = DIR_DOWN Then
    BitBlt picSlider.hdc, 10, iSize + 10, (iSize * 2) + 10, picSlider.ScaleHeight - iSize - 10, picSlider.hdc, 10, 0, vbSrcCopy

    picSlider.Line (0, 10)-(picSlider.Width, iSize + 10), vbWindowBackground, BF

    For nIndex = NUM_OF_PREVIEWS To 3 Step -1
        lPreviewPics(nIndex) = lPreviewPics(nIndex - 2)
    Next nIndex

    lPreviewPics(1) = lPreviewPics(1) - 1
    lPreviewPics(2) = lPreviewPics(2) - 1
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Function FindSlot(ByVal X As Integer, ByVal Y As Integer, ByVal lOld As Long, Optional ByVal bDraw As Byte) As Byte

On Error GoTo ErrorHandler

Dim nIndex As Integer
Dim xSlot As Byte
Dim ySlot As Byte
Dim XPos As Integer
Dim YPos As Integer
Dim iPos As Integer
Dim bNo As Byte

Dim iOffsetX As Integer

iOffsetX = ((picShow.ScaleWidth - picSlider.ScaleWidth) / 2)

Y = Y + iYPos

If X >= 5 + iOffsetX And X <= 10 + iSize + iOffsetX Then
    xSlot = 1
ElseIf X >= 15 + iSize + iOffsetX And X <= picShow.ScaleWidth + iOffsetX - 10 Then
    xSlot = 2
End If

For nIndex = 0 To (NUM_OF_PREVIEWS \ 2) - 1
    If Y >= nIndex * (iSize + 10) + 5 And Y <= nIndex * (iSize + 10) + iSize + 10 Then
        ySlot = nIndex
        Exit For
    ElseIf nIndex = (NUM_OF_PREVIEWS \ 2) - 1 Then
        nIndex = -1
        Exit For
    End If
Next nIndex

iPos = (2 * ySlot) + xSlot

If iPos <> 0 Then If lPreviewPics(iPos) = 0 Or lPreviewPics(iPos) > tPictureFiles.Count Then bNo = 1

If bDraw = 0 Then
    If bNo = 0 And lOld <> (2 * ySlot) + xSlot And nIndex <> -1 And xSlot <> 0 Then
        XPos = (xSlot - 1) * iSize + (10 * xSlot) + iOffsetX
        YPos = ySlot * (iSize + 10) - iYPos + 10
    
        BitBlt picShow.hdc, XPos - 5, YPos - 5, iSize + 5, iSize + 5, picSlider.hdc, XPos - iOffsetX, YPos + iYPos, vbSrcCopy
        picShow.Line (XPos - 5, YPos - 5)-(XPos - 5, YPos + iSize - 5), vbBlack
        picShow.Line (XPos - 5, YPos - 5)-(XPos + iSize - 5, YPos - 5), vbBlack
        picShow.Line (XPos + iSize - 5, YPos - 5)-(XPos + iSize - 5, YPos + iSize - 5), vbBlack
        picShow.Line (XPos - 5, YPos + iSize - 5)-(XPos + iSize - 5, YPos + iSize - 5), vbBlack

        picShow.Line (XPos + iSize - 5, YPos - 3)-(XPos + iSize - 5, YPos + iSize - 5), 0
        picShow.Line (XPos - 3, YPos + iSize - 5)-(XPos + iSize - 4, YPos + iSize - 5), 0
        picShow.Line (XPos + iSize - 4, YPos - 1)-(XPos + iSize - 4, YPos + iSize - 3), RGB(80, 151, 95)
        picShow.Line (XPos - 1, YPos + iSize - 4)-(XPos + iSize - 4, YPos + iSize - 4), RGB(80, 151, 95)
        picShow.Refresh
    End If

    If lOld <> 0 And (lOld <> (2 * ySlot) + xSlot Or nIndex = -1) Then
        If lOld Mod 2 = 0 Then
            ySlot = lOld / 2 - 1
            xSlot = 2
        Else
            ySlot = (lOld + 1) / 2 - 1
            xSlot = 1
        End If

        XPos = ((xSlot - 1) * iSize + (10 * xSlot)) - 6 + iOffsetX
        YPos = (ySlot * (iSize + 10) - iYPos + 10) - 5
        BitBlt picShow.hdc, XPos, YPos, iSize + 5, iSize + 5, picSlider.hdc, XPos - iOffsetX, YPos + iYPos, vbSrcCopy
        picShow.Refresh
    End If
End If

If bNo = 0 And xSlot <> 0 And nIndex <> -1 Then FindSlot = iPos

Exit Function

ErrorHandler:
ErrHandle
Resume Next

End Function
Private Sub InitDirectSound()

' Purpose: Initializes DirectSound and loads sound effects.

On Error GoTo ErrorHandler

If DirectSound.InitDirectSound(Me.hwnd) = True Then
    ' Load sound effects.
    DirectSound.OpenSound "WELCOME", False, CREATE_FROM_RES, 1, "WAV"
    DirectSound.OpenSound "SELECT", False, CREATE_FROM_RES, 2, "WAV"
    DirectSound.OpenSound "HELP", True, CREATE_FROM_RES, 3, "WAV"
End If

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub txtTime_Change()

On Error GoTo ErrorHandler

Dim lColor As Long

If txtTime.Text = "" Then txtTime.Text = 1
If txtTime.Text > 20 Then txtTime.Text = 20
If txtTime.Text = 0 Then txtTime.Text = 1

If iPoint - 1 = txtTime.Text Then Exit Sub

lColor = picDial.ForeColor

picDial.ForeColor = picDial.BackColor

picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

picDial.ForeColor = lColor

picDial.PSet (iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

iPoint = Int(txtTime.Text) - 1

picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

tProgramOptions.iInterval = (iPoint + 1) * 1000

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub txtTime_KeyPress(KeyAscii As Integer)

If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub
Private Sub VScroll_Change()

On Error GoTo ErrorHandler

Dim bIndex As Byte
Dim bTemp As Byte

If bByCode = 1 Then Exit Sub

If Abs(VScroll.Value - lLastVal) = 1 Then
    iYPos = ((VScroll.Value + 1) - (lPreviewPics(2) / 2)) * (iSize + 10)

    If iYPos < 0 Then iYPos = 0

    For bIndex = 1 To Abs(VScroll.Value - lLastVal)
        If iYPos = 0 Then
            UpdatePreview

            If lPreviewPics(1) <> 1 Then
                AdjustRows 1
                bTemp = bNextFree
                bNextFree = 1
                LoadRow DIR_UP, 0

                bNextFree = bTemp + 2

                If bNextFree > NUM_OF_PREVIEWS + 1 Then bNextFree = NUM_OF_PREVIEWS + 1
            End If
        ElseIf HitEdge = True And lPreviewPics(bNextFree - 1) < tPictureFiles.Count Then
            If bNextFree = NUM_OF_PREVIEWS + 1 Then
                picShow.Cls
                bNextFree = bNextFree - 2
                AdjustRows 0
                lPreviewPics(bNextFree) = 0
                lPreviewPics(bNextFree + 1) = 0
                iYPos = (iSize + 10) * (NUM_OF_PREVIEWS / 2 - 2)
                If lPreviewPics(bNextFree) = 0 Then LoadRow DIR_DOWN, 0
            ElseIf lPreviewPics(bNextFree) = 0 Then LoadRow DIR_DOWN, 1
            End If
        Else: UpdatePreview
        End If
    Next bIndex

    UpdatePreview
Else
    picShow.Cls
    picShow.Refresh

    iYPos = 0

    bNextFree = 0

    picSlider.Cls

    For bIndex = 1 To NUM_OF_PREVIEWS
        lPreviewPics(bIndex) = 0
    Next bIndex

    lPreviewPics(1) = (VScroll.Value + 1) * 2 - 2
    LoadRow DIR_DOWN, 1
    LoadRow DIR_DOWN, 1
End If

lLastVal = VScroll.Value

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
Private Sub ContextHelp(ByVal bPlaySound As Byte)

On Error GoTo ErrorHandler

' Context sensitive help.
Select Case bLastClicked
    Case 1: ShowHelp Me.hwnd, "Intro_Introducing.htm", bPlaySound
    Case 2: ShowHelp Me.hwnd, "Options_SelectPictures.htm", bPlaySound
    Case 3: ShowHelp Me.hwnd, "Options_Preview.htm", bPlaySound
    Case 4
        Select Case bCurrentOption
            Case 1: ShowHelp Me.hwnd, "Options_AdjOpt_Interval.htm", bPlaySound
            Case 2: ShowHelp Me.hwnd, "Options_AdjOpt_Transitions.htm", bPlaySound
            Case 3: ShowHelp Me.hwnd, "Options_AdjOpt_Miscellaneous.htm", bPlaySound
            Case 4: ShowHelp Me.hwnd, "Options_AdjOpt_Advanced.htm", bPlaySound
        End Select
    Case 5: ShowHelp Me.hwnd, "Options_BeginScrolling.htm", bPlaySound
End Select

Exit Sub

ErrorHandler:
ErrHandle
Resume Next

End Sub
