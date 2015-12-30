VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   1830
   ClientTop       =   2025
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
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   735
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   11
      Top             =   3780
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
      TabIndex        =   1
      Top             =   5595
      Width           =   7710
   End
   Begin VB.PictureBox picLogo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1215
      Left            =   0
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   0
      Width           =   7710
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   4
      Left            =   3120
      TabIndex        =   5
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   2
         Left            =   165
         TabIndex        =   16
         Top             =   795
         Visible         =   0   'False
         Width           =   4365
         Begin VB.Frame fraMoveDown 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2425
            TabIndex        =   63
            Top             =   2710
            Width           =   1800
            Begin MSComctlLib.Toolbar tlbMoveDown 
               Height          =   270
               Left            =   0
               TabIndex        =   64
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
                  Picture         =   "frmMain.frx":030A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":057C
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
            TabIndex        =   61
            Top             =   715
            Width           =   1800
            Begin MSComctlLib.Toolbar tlbMoveUp 
               Height          =   270
               Left            =   0
               TabIndex        =   62
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
            ItemData        =   "frmMain.frx":07EE
            Left            =   0
            List            =   "frmMain.frx":07F0
            TabIndex        =   59
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
            ItemData        =   "frmMain.frx":07F2
            Left            =   2410
            List            =   "frmMain.frx":07F4
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   1890
            WhatsThisHelpID =   59
            Width           =   540
         End
         Begin VB.CheckBox chkRandom 
            Caption         =   "Tran&sitions Appear Randomly"
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
            TabIndex        =   55
            Top             =   3045
            WhatsThisHelpID =   64
            Width           =   2415
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "<<<"
            Height          =   390
            Left            =   1785
            TabIndex        =   54
            Top             =   2550
            WhatsThisHelpID =   60
            Width           =   540
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   ">>>"
            Height          =   390
            Left            =   1785
            TabIndex        =   53
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   60
            Top             =   105
            Width           =   4125
         End
      End
      Begin MSComctlLib.TabStrip tbsOptions 
         Height          =   3840
         Left            =   105
         TabIndex        =   13
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
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   795
         Visible         =   0   'False
         Width           =   4365
         Begin VB.PictureBox picDial 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            DrawWidth       =   3
            ForeColor       =   &H00000000&
            Height          =   2500
            Left            =   1560
            ScaleHeight     =   165
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   165
            TabIndex        =   83
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
            TabIndex        =   85
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
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
            Left            =   840
            TabIndex        =   84
            Top             =   1680
            Width           =   45
         End
         Begin VB.Label lblInstructions 
            AutoSize        =   -1  'True
            Caption         =   "Use the dial below to select the interval between pictures."
            Height          =   480
            Index           =   4
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   3855
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   4
         Left            =   165
         TabIndex        =   18
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
            TabIndex        =   78
            Top             =   2310
            Width           =   4110
            Begin VB.PictureBox picBackColor 
               Height          =   225
               Left            =   1785
               ScaleHeight     =   165
               ScaleWidth      =   1320
               TabIndex        =   80
               Top             =   315
               Width           =   1380
            End
            Begin VB.PictureBox picInfoColor 
               Height          =   225
               Left            =   1785
               ScaleHeight     =   165
               ScaleWidth      =   1320
               TabIndex        =   79
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
               TabIndex        =   82
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
               TabIndex        =   81
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
            TabIndex        =   47
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
               TabIndex        =   52
               Top             =   945
               Width           =   2640
            End
            Begin VB.OptionButton optScreenSize 
               Caption         =   "Optimize for &Speed"
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
               Top             =   1260
               Width           =   2640
            End
            Begin VB.OptionButton optScreenSize 
               Caption         =   "&Use the current screen settings"
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
               TabIndex        =   48
               Top             =   315
               Width           =   2640
            End
         End
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   3405
         Index           =   3
         Left            =   165
         TabIndex        =   17
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
            TabIndex        =   77
            Top             =   3100
            Width           =   3585
         End
         Begin VB.Frame frmPicture 
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
            TabIndex        =   43
            Top             =   105
            Width           =   2190
            Begin VB.OptionButton optPictureSize 
               Caption         =   "&Original Picture Size"
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
               TabIndex        =   46
               Top             =   315
               WhatsThisHelpID =   69
               Width           =   1845
            End
            Begin VB.OptionButton optPictureSize 
               Caption         =   "&Stretch to fit Screen"
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
               TabIndex        =   45
               Top             =   885
               WhatsThisHelpID =   68
               Width           =   1845
            End
            Begin VB.OptionButton optPictureSize 
               Caption         =   "Stretch &Proportionally"
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
               TabIndex        =   44
               Top             =   600
               WhatsThisHelpID =   69
               Width           =   1845
            End
         End
         Begin VB.Frame frmScroll 
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
            TabIndex        =   39
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
               TabIndex        =   42
               Top             =   315
               WhatsThisHelpID =   65
               Width           =   1125
            End
            Begin VB.OptionButton optDirection 
               Caption         =   "Back&ward"
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
               TabIndex        =   41
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
               TabIndex        =   40
               Top             =   885
               WhatsThisHelpID =   67
               Width           =   1125
            End
         End
         Begin VB.Frame frmMusic 
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
            TabIndex        =   20
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   31
               Top             =   630
               Visible         =   0   'False
               Width           =   3675
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
                  Left            =   210
                  TabIndex        =   38
                  Top             =   105
                  WhatsThisHelpID =   73
                  Width           =   1155
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
                  Left            =   210
                  TabIndex        =   37
                  Top             =   457
                  WhatsThisHelpID =   74
                  Width           =   1155
               End
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
                  TabIndex        =   36
                  Top             =   420
                  WhatsThisHelpID =   74
                  Width           =   1740
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
                     TabIndex        =   34
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
                     TabIndex        =   33
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
                  TabIndex        =   28
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
                  TabIndex        =   27
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
                  TabIndex        =   26
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
                  TabIndex        =   25
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
               TabIndex        =   29
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
                  TabIndex        =   30
                  Top             =   210
                  Width           =   3600
               End
            End
         End
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
         TabIndex        =   14
         Top             =   105
         Width           =   2535
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   5
      Left            =   3150
      TabIndex        =   6
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
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
         Left            =   210
         TabIndex        =   74
         Top             =   2415
         Width           =   4140
      End
      Begin VB.Label lblScrollDirection 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   2100
         TabIndex        =   73
         Top             =   945
         Width           =   75
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   2100
         TabIndex        =   72
         Top             =   525
         Width           =   75
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Direction:"
         Height          =   240
         Index           =   9
         Left            =   210
         TabIndex        =   71
         Top             =   945
         Width           =   1665
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Interval:"
         Height          =   240
         Index           =   8
         Left            =   210
         TabIndex        =   70
         Top             =   525
         Width           =   840
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   3
      Left            =   3084
      TabIndex        =   2
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.PictureBox Picture1 
         Height          =   4110
         Left            =   105
         ScaleHeight     =   4050
         ScaleWidth      =   4365
         TabIndex        =   86
         Top             =   105
         Width           =   4425
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   1
      Left            =   3084
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   210
         Width           =   4335
      End
   End
   Begin VB.Frame fraSteps 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   2
      Left            =   3084
      TabIndex        =   3
      Top             =   1215
      Visible         =   0   'False
      Width           =   4626
      Begin VB.TextBox txtLocation 
         BackColor       =   &H8000000F&
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   3990
         WhatsThisHelpID =   77
         Width           =   2835
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
               Picture         =   "frmMain.frx":07F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0950
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
               Picture         =   "frmMain.frx":0AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0BBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0CCE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0DE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0EF2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   2205
         Left            =   210
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   1050
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   1058
         ButtonWidth     =   1402
         ButtonHeight    =   953
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
               Caption         =   "&Remove"
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
               Caption         =   "Pre&view"
               Key             =   "PREVIEW"
               Object.ToolTipText     =   "Show a preview of the picture"
               ImageIndex      =   5
               Style           =   1
            EndProperty
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
         TabIndex        =   76
         Top             =   4035
         Width           =   1080
      End
      Begin VB.Label lblInstructions 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":1004
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   120
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DT_VCENTER = &H4
Const DT_SINGLELINE = &H20

Const COLOR_BTNFACE = 15
Const COLOR_HIGHLIGHT = 13

Const COLOR_BACK_REG = &HC00000
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
' Used to get a folder from the user

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
' different types of picture files
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_USEFILEATTRIBUTES = &H10
Private Const DI_NORMAL = &H3

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
' Used for drawing the interval dial.

Dim bDown As Byte
Dim iPoint As Integer

Private Type POINTAPI
    X As Integer
    Y As Integer
End Type

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

' Holds the dimensions for each button
Dim rButtons() As RECT
' Specifies the number of buttons
Dim bNumOfOptions As Byte

Dim iCirclePoints(19) As POINTAPI
Private Sub chkMusicType_Click(Index As Integer)

Dim nIndex As Byte

If bByCode = True Then Exit Sub

SetMusicType Index

End Sub
Private Sub chkMusicType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 366 + Index

End Sub
Private Sub chkRandom_Click()

tProgramOptions.bRandomTransitions = chkRandom.Value

UpdateTransButtons

End Sub
Private Sub chkRandom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 348

End Sub
Private Sub chkSoundEffects_Click()

tProgramOptions.bSoundEffects = chkSoundEffects.Value

If tProgramOptions.bSoundEffects = 1 Then
    If DirectSound.InitDirectSound(Me.hwnd) = True Then
        ' Load the necessary welcome tune.
        DirectSound.OpenSound "WELCOME", CREATE_FROM_RES, 1, "WAV"
        DirectSound.OpenSound "BUTTON_OVER", CREATE_FROM_RES, 2, "WAV"
        DirectSound.OpenSound "SELECT", CREATE_FROM_RES, 3, "WAV"
    End If
Else: Set DirectSound = Nothing
End If

End Sub
Private Sub chkSoundEffects_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 377

End Sub
Private Sub cmbDisplayModes_Click()

With tProgramOptions.tDisplayMode
    .iWidth = DirectDraw.ModeWidth(cmbDisplayModes.ListIndex)
    .iHeight = DirectDraw.ModeHeight(cmbDisplayModes.ListIndex)
    .bBPP = DirectDraw.ModeBPP(cmbDisplayModes.ListIndex)
End With

End Sub
Private Sub cmbDisplayModes_GotFocus()

optScreenSize(3).Value = True

End Sub
Private Sub cmbTracks_Click()

tProgramOptions.iTrackNumber = cmbTracks.ListIndex + 1

End Sub
Private Sub cmdAdd_Click()

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

End Sub
Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 342

End Sub
Private Sub cmdAddAll_Click()

lstAvailable.ListIndex = 0

Do Until lstAvailable.ListCount = 0
    cmdAdd_Click
Loop

End Sub
Private Sub cmdAddAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 341

End Sub
Private Sub cmdBrowse_Click()

Dim sSelectedFiles() As String
Dim sFolder As String

' Show the OpenFile CommonDialog box
sFolder = OpenFileDialog(Me.hwnd, MUSIC, sSelectedFiles)

' If the user didn't select a file, exit
If sSelectedFiles(0) = "" Then Exit Sub

tProgramOptions.sMusicFile = sSelectedFiles(0)
txtMusicFile.Text = sSelectedFiles(0)
txtMusicFile.SetFocus

End Sub
Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 376

End Sub
Private Sub cmdCancel_Click()

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

End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 372

End Sub
Private Sub cmdRemove_Click()

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

End Sub
Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 343

End Sub
Private Sub cmdRemoveAll_Click()

lstActual.ListIndex = 0

Do Until lstActual.ListCount = 0
    cmdRemove_Click
Loop

End Sub
Private Sub cmdRemoveAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 344

End Sub
Private Sub cmdRetry_Click()

FillTracks

End Sub
Private Sub cmdRetry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 371

End Sub
Private Sub Form_Click()

Dim hBrush As Long

If bMouseOver > 0 Then
    ' Reset the previous option clicked
    If bLastClicked > 0 And bLastClicked <> bMouseOver Then
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

    ' Play "click" sound if sound effects enabled.
    If tProgramOptions.bSoundEffects = 1 Then DirectSound.PlaySound "SELECT", False

    ' Set which button is selected
    bLastClicked = bMouseOver

    ' First setup the next frame to be shown
    SetupOptions
    fraSteps(bLastClicked).Visible = True
    DoEvents

    ' Draw the background for the button
    hBrush = GetSysColorBrush(COLOR_BTNFACE)
    FillRect Me.hdc, rButtons(bLastClicked), hBrush

    ' Draw the button's caption
    Me.ForeColor = GetSysColor(COLOR_HIGHLIGHT)
    DrawTextAPI Me.hdc, LoadResString(bLastClicked), Len(LoadResString(bLastClicked)), rButtons(bLastClicked), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP

    Me.Refresh
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Dim bClickButton As Byte

If bLastClicked = 2 And (Shift And vbAltMask) Then
    Select Case KeyCode
        Case vbKeyD: tlbDirectory_ButtonClick tlbDirectory.Buttons(1)
        Case vbKeyR: tlbDirectory_ButtonClick tlbDirectory.Buttons(2)
        Case vbKeyO: tlbDirectory_ButtonClick tlbDirectory.Buttons(4)
        Case vbKeyU: tlbDirectory_ButtonClick tlbDirectory.Buttons(5)
        Case vbKeyV
            tlbDirectory.Buttons(7).Value = Abs(tlbDirectory.Buttons(7).Value - 1)
            tlbDirectory_ButtonClick tlbDirectory.Buttons(7)
    End Select
Else
    ' Allow the user to press alt + a key to access
    ' the various button options on the left
    If Shift And vbAltMask Then
        Select Case KeyCode
            Case vbKeyI
                bMouseOver = 1
                bClickButton = True
            Case vbKeyD
                bMouseOver = 2
                bClickButton = True
            Case vbKeyP
                bMouseOver = 3
                bClickButton = True
            Case vbKeyA
                bMouseOver = 4
                bClickButton = True
            Case vbKeyB
                bMouseOver = 5
                bClickButton = True
        End Select

        If bClickButton = True Then Form_Click
    End If
End If

End Sub
Private Sub Form_Load()

Dim rOptions As RECT
Dim hBrush As Long
Dim nIndex As Byte

WaitProcess "Loading Settings", False, Me

Me.Caption = App.Title

' If we're being run in "ScreenSaver Configuration Mode",
' then somethings need to change.
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

' Position the option pages exactly
For nIndex = 1 To 4
    fraOptions(nIndex).Left = tbsOptions.ClientLeft
    fraOptions(nIndex).Top = tbsOptions.ClientTop
    fraOptions(nIndex).Width = tbsOptions.ClientWidth
    fraOptions(nIndex).Height = tbsOptions.ClientHeight
Next nIndex

' Retrieve all the different picture icons
GetPicIcons

' Setup the options section area
With rOptions
    .Left = 0
    .Top = picLogo.Height
    .Right = fraSteps(1).Left
    .bottom = picInfo.Top
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

For nIndex = 1 To bNumOfOptions
    ' Resize all the option frames to the correct size
    With fraSteps(nIndex)
        .Left = Me.ScaleWidth * 0.4
        .Width = Me.ScaleWidth - .Left
        .Top = picLogo.Height
        .Height = picInfo.Top - .Top
    End With

    ' Draw divider lines and the options' text
    With rButtons(nIndex)
        If nIndex <> 1 Then
            .Top = rButtons(nIndex - 1).bottom + 1
        Else: .Top = rOptions.Top
        End If

        .Right = rOptions.Right
        .bottom = .Top + 25
    End With

    Me.ForeColor = COLOR_LINES
    MoveToEx Me.hdc, 0, rButtons(nIndex).bottom, 0
    LineTo Me.hdc, rButtons(nIndex).Right, rButtons(nIndex).bottom

    Me.ForeColor = COLOR_TEXT_REG
    DrawTextAPI Me.hdc, LoadResString(nIndex), Len(LoadResString(nIndex)), rButtons(nIndex), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP
Next nIndex

' Select the first option
bMouseOver = 1
Form_Click

' Set all the options to the user's last.
SetUserOptions

If bRunMode = RM_SAVER_CONFIG Then
    ' Load the list of ScreenSaver pictures
    LoadSavedList sAppPath & "SSList.pcs"
End If

' If we're to play sounds, then play the welcome message.
If tProgramOptions.bSoundEffects = 1 Then
    ' Play the "Welcome Words" :)
    DirectSound.PlaySound "WELCOME", False
End If

EndWaitProcess Me

Me.Show
DoEvents

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim nIndex As Byte
Dim bInButton As Byte

For nIndex = 1 To bNumOfOptions
    With rButtons(nIndex)
        ' See if the mouse is within the area of this
        ' particular button
        If X > .Left And X < .Right And Y > .Top And Y < .bottom Then
            ' Just exit if this is the same button we
            ' were on before.
            If nIndex = bMouseOver Then Exit Sub

            ' Play "over" sound if sound effects enabled.
            If tProgramOptions.bSoundEffects = 1 Then DirectSound.PlaySound "BUTTON_OVER", False

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

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

bCancelOp = True

tProgramOptions.bNotFirstStart = True

' Save the user's options
Open sAppPath & "Options.dat" For Binary As #1
Put #1, , tProgramOptions
Close #1

If bRunMode = RM_SAVER_CONFIG Then SaveCurrentList sAppPath & "SSList.pcs"

Set DirectSound = Nothing

End Sub
Private Sub fraSteps_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 5 Then
    With lblBeginScrolling
        .FontUnderline = False
        .ForeColor = vbButtonText
    End With
End If

End Sub
Private Sub lblBeginScrolling_Click()

If tPictureFiles.Count = 0 Then Exit Sub

If tProgramOptions.bNotFirstStart = False Then MsgBox "In order to show/hide the scrolling controls, press the 'H' key." & vbCr & vbCr & "This message will only be shown once.", vbOKOnly + vbInformation

Me.Hide
DoEvents

Load frmScroller
DoEvents

Unload frmScroller
DoEvents

Me.Show
DoEvents

End Sub
Private Sub lblBeginScrolling_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

With lblBeginScrolling
    .FontUnderline = True
    .ForeColor = vbBlue
End With

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

Dim sPath As String

If lstFiles.SelectedItem.Key = "NO_FILES" Then Exit Sub

If lstFiles.SelectedItem.Key = "UP_LEVEL" Then
    AddFilesToList ""
Else
    ' Figure out the path of the selected file
    sPath = GetItemPath(lstFiles.SelectedItem.Index)

    If GetAttr(sPath) And vbDirectory Then AddFilesToList sPath
End If

End Sub
Private Sub lstFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim sPath As String
Dim lPictureIndex As Long
Dim rPreview As RECT

UpdateDirButtons

' Exit if it's not applicable to load the preview.
If tlbDirectory.Buttons("PREVIEW").Value = tbrUnpressed Or _
    Item.Key = "NO_FILES" Or Item.Key = "UP_LEVEL" Or _
    sCurrentFolder = "" Then GoTo NotPicture

' Extract the picture index from the key of the item.
lPictureIndex = Right(Item.Key, Len(Item.Key) - 4)

' Attempt to load the picture.
If GetPicFromIndex(picHidden, lPictureIndex) = False Then
    If MsgBox("This file doesn't seem to be a valid picture.  Would you like to remove it from the list?", vbYesNo + vbExclamation) = vbYes Then
        picPreview.Picture = LoadPicture()

        lstFiles.ListItems.Remove Item.Index

        AddFilesToList lstFiles.SelectedItem.Text
    Else
        With rPreview
            .Right = picPreview.Width
            .bottom = picPreview.Height
        End With

        picPreview.Picture = LoadPicture()
        DrawTextAPI picPreview.hdc, "Invalid Picture", 15, rPreview, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOCLIP
    End If
Else
    ' If we have successfully loaded the picture, then
    ' put the picture in the picture box.
    If tProgramOptions.bPreviewSize = 0 Then
        StretchBlt picPreview.hdc, 0, 0, picPreview.Width, picPreview.Height, picHidden.hdc, 0, 0, picHidden.Width, picHidden.Height, vbSrcCopy
    Else: BitBlt picPreview.hdc, 0, 0, picPreview.Width, picPreview.Height, picHidden.hdc, 0, 0, vbSrcCopy
    End If

    picPreview.Refresh
End If

Exit Sub

NotPicture:
picPreview.Picture = LoadPicture()

End Sub
Private Sub lstFiles_KeyDown(KeyCode As Integer, Shift As Integer)

Dim nIndex As Long

' Allow the user to press Ctrl+A to select all the files
If KeyCode = vbKeyA And (Shift And vbCtrlMask) > 0 Then
    For nIndex = 1 To lstFiles.ListItems.Count
        lstFiles.ListItems(nIndex).Selected = True
    Next nIndex
ElseIf KeyCode = vbKeyDelete Then
    tlbDirectory_ButtonClick tlbDirectory.Buttons(2)
End If

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

DrawTip 373 + Index

End Sub
Private Sub optPictureSize_Click(Index As Integer)

tProgramOptions.bPictureSize = Index

End Sub
Private Sub optPictureSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 363 + Index

End Sub
Private Sub optPlayTrack_Click(Index As Integer)

tProgramOptions.bPlayTrack = Index
cmbTracks.Enabled = Index

End Sub
Private Sub optPlayTrack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 369 + Index

End Sub
Private Sub optScreenSize_Click(Index As Integer)

tProgramOptions.bScreenSetting = Index

End Sub
Private Sub optScreenSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 380 + Index

End Sub
Private Sub picBackColor_Click()

Dim lNewColor As Long

lNewColor = SelectColor(Me.hwnd, tProgramOptions.lBackColor)

If lNewColor <> -1 Then
    tProgramOptions.lBackColor = lNewColor
    picBackColor.BackColor = lNewColor
End If

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

Dim nIndex As Integer
Dim iShortest As Integer
Dim dTemp As Double
Dim dShortest As Double

If bDown = 1 Then
    picDial.ForeColor = vbWhite

    picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    picDial.ForeColor = vbBlack

    picDial.PSet (iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    For nIndex = 0 To 19
        dTemp = Sqr((iCirclePoints(nIndex).X - X) ^ 2 + (iCirclePoints(nIndex).Y - Y) ^ 2)

        If dTemp < dShortest Or dShortest = 0 Then
            dShortest = dTemp
            iShortest = nIndex
        End If
    Next nIndex

    iPoint = iShortest

    lblTime.Caption = iPoint + 1

    picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    'tProgramOptions.iInterval = (iPoint + 1) * 1000
End If

End Sub
Private Sub picDial_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

bDown = 0

End Sub
Private Sub picInfoColor_Click()

Dim lNewColor As Long

lNewColor = SelectColor(Me.hwnd, tProgramOptions.lInfoColor)

If lNewColor <> -1 Then
    tProgramOptions.lInfoColor = lNewColor
    picInfoColor.BackColor = lNewColor
End If

End Sub
Private Sub picInfoColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 385

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
Private Sub picPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Show the picture stretch menu when they right click.
If Button = 2 Then
    frmMenu.mnuStretchItem(tProgramOptions.bPreviewSize).Checked = True
    ShowPopMenu frmMenu.mnuStretch, picPreview.Left + X, picPreview.Top + Y
End If

End Sub
Private Sub tbsOptions_Click()

Dim bLastOption As Byte

bLastOption = bCurrentOption

' Show the the frame of the option clicked.
bCurrentOption = tbsOptions.SelectedItem.Index
fraOptions(bCurrentOption).Visible = True
DoEvents

' We do this here instead of at load time, because
' we don't want to bother the user until they do there.
If bCurrentOption = 3 Then SetMusicType tProgramOptions.bMusicType

' Hide the previous option frame.
fraOptions(bLastOption).Visible = False
DoEvents

End Sub
Private Sub tbsOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim nIndex As Byte

For nIndex = 1 To tbsOptions.Tabs.Count
    With tbsOptions.Tabs(nIndex)
        If X > .Left And X < .Left + .Width And (Y + tbsOptions.Top) > .Top And (Y + tbsOptions.Top) < .Top + .Height Then
            DrawTip 299 + nIndex
        End If
    End With
Next nIndex

End Sub
Private Sub tlbDirectory_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim iToolbarX As Integer
Dim iToolbarY As Integer

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

End Sub
Private Sub SetupOptions()

' Purpose: Setup various settings depending on the
'   option selected (the buttons on the left).

Select Case bLastClicked
    Case 2
        ' Add files to the list if necessary
        If lstFiles.ListItems.Count = 0 Then AddFilesToList sCurrentFolder

        ' If the user has previously shown the preview
        ' button, then show it again.
        If tlbDirectory.Buttons("PREVIEW").Value = tbrPressed Then picPreview.Visible = True
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
End Select

' Hide the preview window if necessary
If bMouseOver <> 2 Then
    If picPreview.Visible = True Then picPreview.Visible = False
End If

End Sub
Private Sub SearchForPics(ByVal sPath As String, ByVal bRecursive As Byte)

' Purpose: Searchs through the specified directory
'   and adds only picture files to the array
'   "tPictureFiles". This function is called by itself
'   in order to search subfolders.

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
        sResult = Left(tFileData.cFileName, iPosition - 1)
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

End Sub
Private Function GetFolder() As String

' Purpose: Show the "Browser For Folder" window
'   and return the result

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
    sPath = String(255, 0)

    ' Retrieve a "true" path from the PIDL number
    SHGetPathFromIDList lFolderPIDL, sPath

    iPosition = InStr(sPath, vbNullChar)

    ' Remove the last null character if there is one
    ' and set "sPath" to that value
    If iPosition <> 0 Then
        GetFolder = Left(sPath, iPosition - 1)
    Else
        GetFolder = sPath
    End If
End If

End Function
Private Sub AddFilesToList(ByVal sFolder As String)

' Purpose: Add the correct information to the listbox

Dim nIndex As Long
Dim lFolderID As Long
Dim sPath As String
Dim sTemp As String
Dim iLastSlash As Integer

Static bAddingItems As Byte

If bAddingItems = True Then Exit Sub

bAddingItems = True

Screen.MousePointer = 11

' Set the folder that is currently shown ("" if
' all the folders are in view).
sCurrentFolder = sFolder

' Put that in a format the user can understand.
If sCurrentFolder = "" Then
    txtLocation.Text = " All Folders"
Else: txtLocation.Text = " " & sCurrentFolder
End If

lstFiles.ListItems.Clear

UpdateDirButtons

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
        If tPictureFiles(nIndex).FolderID = lFolderID Then
            sPath = tPictureFiles(nIndex).FileName

            iLastSlash = LastSlash(sPath)

            lstFiles.ListItems.Add , "PIC_" & nIndex, Right(sPath, Len(sPath) - iLastSlash), , tPictureFiles(nIndex).PicType + 2

            ' If the folder is a root drive, then
            ' don't remove the slash
            sTemp = Left(sPath, iLastSlash)
            If Len(sTemp) <> 3 Then sTemp = Left(sTemp, Len(sTemp) - 1)

            lstFiles.ListItems("PIC_" & nIndex).SubItems(1) = Round(FileLen(sPath) \ 1000) & "kb"
            lstFiles.ListItems("PIC_" & nIndex).bOld = True
        End If

        If bCancelOp = True Then
            Exit Sub
        Else: DoEvents
        End If
    Next nIndex
Else
    lstFiles.ColumnHeaders(1).Text = "Folder Name"
    lstFiles.ColumnHeaders(1).Width = 2000
    lstFiles.ColumnHeaders(2).Text = "Parent"
    lstFiles.ColumnHeaders(2).Width = 2000

    For nIndex = 1 To tAddedFolders.Count
        sPath = tAddedFolders(nIndex).FolderName

        ' Remove the last slash from the path only if
        ' it isn't a root drive
        If Len(sPath) <> 3 Then sPath = Left(sPath, Len(sPath) - 1)

        ' If the path is a root drive, then
        ' only put the path in the left column; if
        ' it refers to a directory then put the
        ' folder name in the left side and it's
        ' parent's path in the second column
        If Len(sPath) <> 3 Then
            iLastSlash = LastSlash(sPath)

            lstFiles.ListItems.Add , "DIR_" & tAddedFolders(nIndex).FolderID, Right(sPath, Len(sPath) - iLastSlash), , 1

            ' Get the parent path, but remove the last
            ' slash if it isn't a root drive
            sPath = Left(sPath, iLastSlash)
            If Len(sPath) <> 3 Then sPath = Left(sPath, Len(sPath) - 1)

            lstFiles.ListItems(lstFiles.ListItems.Count).SubItems(1) = sPath
        Else: lstFiles.ListItems.Add , "DIR_" & tAddedFolders(nIndex).FolderID, sPath, , 1
        End If

        lstFiles.ListItems(lstFiles.ListItems.Count).bOld = True

        If bCancelOp = True Then
            Exit Sub
        Else: DoEvents
        End If
    Next nIndex
End If

If lstFiles.ListItems.Count = 0 Then lstFiles.ListItems.Add , "NO_FILES", "No files.", , 2

UpdateDirButtons

Screen.MousePointer = 0

bAddingItems = False

End Sub
Private Sub GetPicIcons()

' Purpose: Load the system icons of different pictures

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

End Sub
Private Sub MovePicture(ByVal sDirection As String)

' Purpose: Flips two items in the collection of pictures.

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

End Sub
Private Sub MoveFolder(ByVal sDirection As String)

' Purpose: Move a folder of pictures up or down through
'   the collection of pictures; it also changes the
'   collection of folders.

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

End Sub
Private Sub FlipItems(ByVal lFirstItem As Long, ByVal lSecondItem As Long, ByVal bDoCollection As Byte)

' Purpose: Flip two items in the lstFiles.  Optionally
'   will also flip the two items in the picture collection

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

    sTempKey1 = Right(sTempKey1, Len(sTempKey1) - 4)
    sTempKey2 = Right(sTempKey2, Len(sTempKey2) - 4)

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
End If

End Sub
Private Function GetItemPath(ByVal lItemIndex As Long) As String

' Purpose: Figure out the path of an item in the listbox
'   This is only meant to be used on directory items

Dim sItemKey As String
Dim lItemID As Long
Dim sType As String

' Get the item's key.
sItemKey = lstFiles.ListItems(lItemIndex).Key
' Extract the ID of the item, either to a file or a folder.
lItemID = Right(sItemKey, Len(sItemKey) - 4)
' Find out what the ID refers to.
sType = Left(sItemKey, 4)

' Retrieve the file name of the specified item.
Select Case sType
    Case "PIC_"
        GetItemPath = tPictureFiles(lItemID).FileName
    Case "DIR_"
        GetItemPath = tAddedFolders("DIR_" & lItemID).FolderName
End Select

End Function
Private Sub SetMusicType(ByVal bType As Byte)

' Purpose: Setup the music type the user selected

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

End Sub
Private Sub FillTracks()

' Purpose: See if a CD is in the drive and if so,
'   add the tracks to cmbTracks

Dim nIndex As Integer

' See if there's a CD is the drive
If BackMusic.GetTrackNumber = True Then
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
Else
    fraError.ZOrder
    fraError.Visible = True
End If

End Sub
Private Sub SetUserOptions()

' Purpose: Sets the value of most of the controls
'   to the user's last used value.

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
    ' Set default options.
    If .bNotFirstStart = False Then
        .iInterval = 1000
        .lInfoColor = RGB(255, 255, 255)
        .bSoundEffects = 1

        ' Default Transitions (all of 'em)!!
        If .bNotFirstStart = False Then
            ReDim .bTransitions(1 To NUM_OF_TRANSITIONS)

            For nIndex = 1 To NUM_OF_TRANSITIONS
                .bTransitions(nIndex) = nIndex
            Next nIndex
        End If

        ' If this is the first start, then load the
        ' example pictures for the user.
        LoadSavedList sAppPath & "SSList.pcs"
    End If

    ReDim .bCustomColors(0 To 16 * 4 - 1) As Byte

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

    picDial.Line (picDial.ScaleWidth \ 2, picDial.ScaleHeight \ 2)-(iCirclePoints(iPoint).X, iCirclePoints(iPoint).Y)

    ' ---------------------------------------------------
    ' Transitions Tab

    ' See if the user has previously set the
    ' transitions to use.
    On Error Resume Next
    If UBound(.bTransitions) = 0 Then
        ' If an error occurs this will also be set.
        bNoTransitions = True
    End If
    On Error GoTo 0

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
    optPictureSize(.bPictureSize).Value = True

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
    For nIndex = 0 To DirectDraw.ModeCount
        cmbDisplayModes.AddItem DirectDraw.ModeWidth(nIndex) & " x " & DirectDraw.ModeHeight(nIndex) & " x " & DirectDraw.ModeBPP(nIndex)
    Next nIndex

    ' If the user has previously set a display
    ' mode to use, then find it and set the
    ' selected item to be it.
    With .tDisplayMode
        If .iWidth <> 0 And .iHeight <> 0 And .bBPP <> 0 Then
            For nIndex = 0 To DirectDraw.ModeCount
                If .iWidth = DirectDraw.ModeWidth(nIndex) And .iHeight = DirectDraw.ModeHeight(nIndex) And .bBPP = DirectDraw.ModeBPP(nIndex) Then
                    cmbDisplayModes.ListIndex = nIndex

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

End Sub
Private Sub DrawTip(ByVal lTipID As Long)

' Purpose: Receives a string of text to print
'   into picInfo in the exact position.

Dim sTip As String
Dim rInfo As RECT
Dim fHeight As Single

Const MARGIN = 5

sTip = LoadResString(lTipID)

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
    .bottom = .Top + fHeight
End With

picInfo.Cls

' Draw the tip for the button
DrawTextAPI picInfo.hdc, sTip, Len(sTip), rInfo, DT_CENTER Or DT_WORDBREAK Or DT_NOCLIP

End Sub
Private Function LookUpTrans(ByVal bTransition As Byte) As String

' Purpose: Give it a transition number and it
'   returns its description text.

Select Case bTransition
    Case 1: LookUpTrans = "Blinds Horizontal"
    Case 2: LookUpTrans = "Blinds Vertical"
    Case 3: LookUpTrans = "Box In"
    Case 4: LookUpTrans = "Box Out"
    Case 5: LookUpTrans = "Smear"
    Case 6: LookUpTrans = "Slide"
    Case 7: LookUpTrans = "Move Up"
    Case 8: LookUpTrans = "Move Down"
    Case 9: LookUpTrans = "Move Left"
    Case 10: LookUpTrans = "Move Right"
End Select

End Function
Private Sub tlbDirectory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim nIndex As Integer

For nIndex = 1 To tlbDirectory.Buttons.Count
    If X >= tlbDirectory.Buttons(nIndex).Left _
        And X <= tlbDirectory.Buttons(nIndex).Left + tlbDirectory.ButtonWidth _
        And Y >= tlbDirectory.Buttons(nIndex).Top _
        And Y <= tlbDirectory.Buttons(nIndex).Top + tlbDirectory.ButtonHeight _
        And nIndex <> 3 And nIndex <> 6 Then
            DrawTip 119 + nIndex
            Exit For
    End If
Next nIndex

End Sub
Private Sub tlbMoveDown_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim sName As String
Dim sData As String

' Disable the DOWN button if moving this one down
' will make it the last item.
If lstActual.List(lstActual.ListIndex + 2) = "" Then tlbMoveDown.Buttons(1).Enabled = False

' Enable the UP button if it is currently disabled.
If tlbMoveUp.Buttons(1).Enabled = False Then tlbMoveUp.Buttons(1).Enabled = True

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

End Sub
Private Sub tlbMoveDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 347

End Sub
Private Sub tlbMoveUp_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim sName As String
Dim sData As String

' Disable the UP button if moving this one up
' will make it the first item.
If lstActual.List(lstActual.ListIndex - 2) = "" Then tlbMoveUp.Buttons(1).Enabled = False

' Enable the DOWN button if it is currently disabled.
If tlbMoveDown.Buttons(1).Enabled <> True Then tlbMoveDown.Buttons(1).Enabled = True

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

End Sub
Private Sub tlbMoveUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 346

End Sub
Private Sub UpdateTransButtons()

' Purpose: On the transitions tab, updates the
'   various buttons to reflex the status of the
'   listboxes and the user's selection.

If lstActual.ListIndex > 0 And tProgramOptions.bRandomTransitions = False Then
    tlbMoveUp.Buttons(1).Enabled = True
Else: tlbMoveUp.Buttons(1).Enabled = False
End If

If lstActual.ListIndex < lstActual.ListCount - 1 And lstActual.ListIndex <> -1 And tProgramOptions.bRandomTransitions = False Then
    tlbMoveDown.Buttons(1).Enabled = True
Else: tlbMoveDown.Buttons(1).Enabled = False
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

End Sub
Private Sub UpdateTransitions()

' Purpose: Recreates the user's settings for the
'   transitions to use.

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

End Sub
Private Sub UpdateDirButtons()

' Purpose: On the "Select Directory" option, updates the
'   various buttons to reflex the status of the listbox
'   and the user's selection.

' If there aren't any items just disable both buttons.
If lstFiles.ListItems.Count = 0 Then GoTo DisableBoth

If lstFiles.ListItems(1).Key = "NO_FILES" Then
    tlbDirectory.Buttons("REMOVE").Enabled = False
Else: tlbDirectory.Buttons("REMOVE").Enabled = True
End If

' If a file is selected but not more than one.
If Not lstFiles.SelectedItem Is Nothing Then
    ' Depending on the view (folders or files), we
    ' can allow the item selected to be moved.
    ' If folders, then from the first item in the list can
    ' they be moved down; otherwise, only from the second
    ' item can they be moved.
    If lstFiles.SelectedItem.Index > IIf(sCurrentFolder = "", 0, 1) And lstFiles.SelectedItem.Index < lstFiles.ListItems.Count Then
        tlbDirectory.Buttons(4).Enabled = True
    Else: tlbDirectory.Buttons(4).Enabled = False
    End If

    ' Depending on the view (folders or files), we
    ' can allow the item selected to be moved.
    ' If folders, then from the second item in the list
    ' can they be moved up; otherwise, only from the
        ' third item they be moved.
    If lstFiles.SelectedItem.Index > IIf(sCurrentFolder = "", 1, 2) Then
        tlbDirectory.Buttons(5).Enabled = True
    Else: tlbDirectory.Buttons(5).Enabled = False
    End If
Else: GoTo DisableBoth
End If

Exit Sub

DisableBoth:
' Disable the buttons if more than one item is
' selected, if no item is selected, or if there aren't
' any items in the list box.
tlbDirectory.Buttons(4).Enabled = False
tlbDirectory.Buttons(5).Enabled = False

End Sub
Private Sub RemoveItems()

' Purpose: Remove either the selected pictures from
'   the collection or a whole folder of pictures.

Dim nIndexDirs As Long
Dim nIndexFiles As Long
Dim lFolderID As Long

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
                tPictureFiles.Remove CLng(Right(lstFiles.ListItems(nIndexFiles).Key, Len(lstFiles.ListItems(nIndexFiles).Key) - 4))
            End If

            lstFiles.ListItems.Remove nIndexFiles
        End If
    Next nIndexFiles

    picPreview.Picture = LoadPicture()

    ' Remove the folder as well if there aren't
    ' any files left
    If lstFiles.ListItems.Count = 0 Then
        tAddedFolders.Remove "DIR_" & GetFolderID(sCurrentFolder)

        AddFilesToList ""
    Else: AddFilesToList sCurrentFolder
    End If

    Screen.MousePointer = 0
End If

End Sub
Private Sub txtMusicFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

DrawTip 375

End Sub
Private Sub txtMusicFile_Validate(Cancel As Boolean)

If Dir(txtMusicFile.Text) = "" Then
    If MsgBox("Please enter a valid path to a music file.", vbOKCancel + vbExclamation) = vbCancel Then
        txtMusicFile.Text = tProgramOptions.sMusicFile
    End If

    Cancel = True
Else: tProgramOptions.sMusicFile = txtMusicFile.Text
End If

End Sub
Public Sub MenuSelect(ByVal hwnd As Long)

' Purpose: This gets called whenever the user's mouse
'   goes over a menu item.

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

End Sub
Private Sub WalkSubMenu(hSubMenu As Long)

' Purpose: Used to "walk" through a sub menu to
'   find the one selected.

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
            sMenuCaption = Left(sMenuCaption, lReturnVal)

            Select Case sMenuCaption
                Case "&Directory"
                    DrawTip 50
                Case "&File(s)"
                    DrawTip 51
                Case "&Load Saved List"
                    DrawTip 52
                Case "&Save Current List"
                    DrawTip 53
                Case "&Stretch"
                    DrawTip 54
                Case "&Normal"
                    DrawTip 55
            End Select

            Exit Sub
        End If
    End If
Next nIndex

End Sub
Private Function AnyLit(hSubMenu As Long) As Byte

' Purpose: Search a menu to see if any items are selected.

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

End Function
Private Sub ShowPopMenu(ByVal objMenu As Menu, ByVal X As Integer, ByVal Y As Integer)

' Purpose: Show a popup menu.  Since the menu will be
'   from frmMenu, unload frmMenu when the popup menu
'   is done.  This will stop subclassing.

Me.PopupMenu objMenu, , X, Y
Unload frmMenu

End Sub
Public Sub AddMenuClick(ByVal iItemIndex As Byte)

' Purpose: Handles when the user clicks an item in
'   the "Add" menu.

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

        bRecursive = MsgBox("Would you like to include the pictures that may be in any sub folders as well?", vbYesNoCancel)

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

        AddFilesToList sFolder

        Screen.MousePointer = 0

    ' If the user clicked load...
    Case 3
GetFileName:
        ' Ask the user for a file.
        OpenFileDialog Me.hwnd, SAVED_LIST, sSelectedFiles

        ' If they didn't select a file, exit.
        If sSelectedFiles(0) = "" Then Exit Sub

        WaitProcess "Loading List", True, Me

        ' Attempt to load the saved list.
        If LoadSavedList(sSelectedFiles(0)) = False Then
            MsgBox "This file doesn't seem to be in the correct format for loading into Picture Scroller." & vbCr & vbCr & "Please select a different file.", vbInformation
            GoTo GetFileName
        End If

        EndWaitProcess Me

        AddFilesToList ""

    ' If the user clicked save...
    Case 4
        sFolder = SaveFileDialog(Me.hwnd)

        If sFolder <> "" Then
            SaveCurrentList sFolder
        End If
End Select

End Sub
Public Sub StretchMenuClick(ByVal nIndex As Byte)

' Purpose: Handles when the user clicks an item in
'   the "Stretch" menu.

tProgramOptions.bPreviewSize = nIndex

End Sub
Private Sub SaveCurrentList(ByVal sFileName As String)

Dim nIndex As Long

Open sFileName For Output As #1

Print #1, LIST_HEADER

For nIndex = 1 To tPictureFiles.Count
    Print #1, tPictureFiles(nIndex).FileName
Next nIndex

Close #1

End Sub
