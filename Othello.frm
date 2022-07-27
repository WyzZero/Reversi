VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Othello 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Othello"
   ClientHeight    =   10950
   ClientLeft      =   3420
   ClientTop       =   870
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   Picture         =   "Othello.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Min 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   26
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Tweak 
      Height          =   135
      Left            =   10800
      TabIndex        =   25
      Top             =   6720
      Width           =   75
   End
   Begin VB.CommandButton Closebtn 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9960
      TabIndex        =   22
      Top             =   615
      Width           =   255
   End
   Begin VB.PictureBox GridRow 
      Enabled         =   0   'False
      Height          =   5835
      Left            =   3000
      ScaleHeight     =   5775
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   1920
      Width           =   315
      Begin VB.CommandButton garbage 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   0
         TabIndex        =   17
         Top             =   5040
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   0
         TabIndex        =   16
         Top             =   4320
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   0
         TabIndex        =   0
         Top             =   3600
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   0
         TabIndex        =   15
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   0
         TabIndex        =   14
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   0
         TabIndex        =   13
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton garbage 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox GridCol 
      Enabled         =   0   'False
      Height          =   310
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   1605
      Width           =   5835
      Begin VB.CommandButton garbage 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   7
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton garbage 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Image Musicbtn 
      Height          =   900
      Left            =   360
      Top             =   5400
      Width           =   2340
   End
   Begin VB.Image Optionbtn 
      Height          =   900
      Left            =   360
      Top             =   4200
      Width           =   2340
   End
   Begin VB.Image Restartbtn 
      Height          =   900
      Left            =   360
      Top             =   3000
      Width           =   2340
   End
   Begin VB.Image Histbtn 
      Height          =   900
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   2340
   End
   Begin VB.Label Bluelbl 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   5880
      TabIndex        =   24
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label Redlbl 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4800
      TabIndex        =   23
      Top             =   9360
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   5760
      Y1              =   8760
      Y2              =   10320
   End
   Begin VB.Label P1namelbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3480
      TabIndex        =   21
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Label P2namelbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   5880
      TabIndex        =   20
      Top             =   8760
      Width           =   2295
   End
   Begin MediaPlayerCtl.MediaPlayer Music 
      Height          =   1455
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Turnlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Turn"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2880
      TabIndex        =   18
      Top             =   7920
      Width           =   5775
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   64
      Left            =   8400
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   63
      Left            =   7680
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   62
      Left            =   6960
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   61
      Left            =   6240
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   60
      Left            =   5520
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   59
      Left            =   4800
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   58
      Left            =   4080
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   57
      Left            =   3360
      Tag             =   "0"
      Top             =   6960
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   56
      Left            =   8400
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   55
      Left            =   7680
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   54
      Left            =   6960
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   53
      Left            =   6240
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   52
      Left            =   5520
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   51
      Left            =   4800
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   50
      Left            =   4080
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   49
      Left            =   3360
      Tag             =   "0"
      Top             =   6240
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   48
      Left            =   8400
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   47
      Left            =   7680
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   46
      Left            =   6960
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   45
      Left            =   6240
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   44
      Left            =   5520
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   43
      Left            =   4800
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   42
      Left            =   4080
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   41
      Left            =   3360
      Tag             =   "0"
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   40
      Left            =   8400
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   39
      Left            =   7680
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   38
      Left            =   6960
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   37
      Left            =   6240
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   36
      Left            =   5520
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   35
      Left            =   4800
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   34
      Left            =   4080
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   33
      Left            =   3360
      Tag             =   "0"
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   32
      Left            =   8400
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   31
      Left            =   7680
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   30
      Left            =   6960
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   29
      Left            =   6240
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   28
      Left            =   5520
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   27
      Left            =   4800
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   26
      Left            =   4080
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   25
      Left            =   3360
      Tag             =   "0"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   24
      Left            =   8400
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   23
      Left            =   7680
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   22
      Left            =   6960
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   21
      Left            =   6240
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   20
      Left            =   5520
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   19
      Left            =   4800
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   18
      Left            =   4080
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   17
      Left            =   3360
      Tag             =   "0"
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   16
      Left            =   8400
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   15
      Left            =   7680
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   14
      Left            =   6960
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   13
      Left            =   6240
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   12
      Left            =   5520
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   11
      Left            =   4800
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   10
      Left            =   4080
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   9
      Left            =   3360
      Tag             =   "0"
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   8
      Left            =   8400
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   7
      Left            =   7680
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   6
      Left            =   6960
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   5
      Left            =   6240
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   4
      Left            =   5520
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   3
      Left            =   4800
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   2
      Left            =   4080
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image Tile 
      Height          =   750
      Index           =   1
      Left            =   3360
      Tag             =   "0"
      Top             =   1920
      Width           =   750
   End
End
Attribute VB_Name = "Othello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------Used for Transparent Forms----------
Dim rgnBasic As New Region
Dim rgnExtended As New Region
Dim CurrentRgn As Long
Dim pic(0 To 1) As New StdPicture
'----------Used for Transparent Forms----------
Dim x As Single



Private Sub Closebtn_Click()
Unload Me
End Sub


Private Sub Form_Load()
Othello.Hide
'----------Used for Transparent Forms----------
    ' Load the image
    Set pic(1) = LoadPicture(App.Path & "/Graphx/Border.gif", 0, 0, 0, 0)
    
    ' Scan the image
    Call rgnBasic.ScanPicture(pic(1))
    
    ' Offset the Shape to allow for the form header.
    Call rgnBasic.OffsetHeader(Me)
        
    Me.picture = pic(1) ' Set the Form Background
    Call rgnBasic.ApplyRgn(Me.hWnd) ' Set the Form Shape
    CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
'----------Used for Transparent Forms----------


Histbtn = LoadPicture(App.Path & "/Graphx/History.gif")
Optionbtn = LoadPicture(App.Path & "/Graphx/Options.gif")
Restartbtn = LoadPicture(App.Path & "/Graphx/Restart.gif")
Musicbtn = LoadPicture(App.Path & "/Graphx/Music.gif")

musicon = True
moveson = True

P1namelbl.Caption = "Player1"
P2namelbl.Caption = "Player2"

song = 1
Musicsel

moved = False
setup

Othello.Show

Log1.Show



End Sub

Sub setup()

'Create the board
For x = 1 To 64
    Tile(x) = LoadPicture(App.Path & "/Graphx/Tile.gif")
    Tile(x).Tag = 0
Next
'----------------

For y = 0 To 7
    For x = 1 To 8
        grid(x, y) = 0
    Next x
Next y

'Get into position
Tile(28) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
Tile(28).Tag = 1
grid(4, 3) = 1

Tile(29) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
Tile(29).Tag = 2
grid(5, 3) = 2

Tile(36) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
Tile(36).Tag = 2
grid(4, 4) = 2

Tile(37) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
Tile(37).Tag = 1
grid(5, 4) = 1
'-----------------

RScore = 2
BScore = 2

Turn = 1
NT = 2 'Not turn

m = 0

OrigX = 0
OrigY = 0
NewX = 0
NewY = 0
repos = 0


Score
Where

Log1.History.Text = ""
Log2.Player1.Text = ""
Log2.Player2.Text = ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
repos = True
OrigX = x
OrigY = y
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If repos = True Then
    moved = True
End If
Tweak.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
repos = False
If moved = True Then
    NewX = x
    NewY = y
    Movedform
End If
Tweak.SetFocus
End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload Log1
Unload Log2
Unload Options
Unload Musicfrm
End Sub

Private Sub Histbtn_Click()
If Log1.Visible = False And Log2.Visible = False Then
    Log1.Visible = True
    Log1.Stylebox.ListIndex = 0
End If

End Sub

Private Sub Min_Click()
Unload Musicfrm
Unload Options
Log1.WindowState = 1
Log1.Visible = True
End Sub

Private Sub Music_EndOfStream(ByVal Result As Long)

newsong:

Dim songr As Integer
Randomize
songr = Int(Rnd * 6)
If songr <> song Then
    song = songr
    Musicsel
Else
    
    GoTo newsong
End If
End Sub

Private Sub New_Click()
setup
End Sub


Private Sub Musicbtn_Click()
Musicfrm.Visible = True
Musicfrm.Musicbox.ListIndex = song
End Sub

Private Sub Optionbtn_Click()
Options.Show

End Sub

Private Sub Restartbtn_Click()
setup
End Sub

Private Sub Tile_Click(Index As Integer)


If Tile(Index).Tag = 4 Then
    Tile(Index).Tag = Turn
    EmptyBox
    pos = Index
    m = m + 1
    hist
    hist2
    If Turn = 1 Then
        Tile(Index) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
        TakeN
        TakeS
        TakeW
        TakeE
        TakeNW
        TakeNE
        TakeSW
        TakeSE
        Turn = 2
        NT = 1
    Else
        Tile(Index) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
        TakeN
        TakeS
        TakeW
        TakeE
        TakeNW
        TakeNE
        TakeSW
        TakeSE
        Turn = 1
        NT = 2
    End If
    Score
    Where
End If
End Sub

Private Sub Tile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Tweak.SetFocus

End Sub
