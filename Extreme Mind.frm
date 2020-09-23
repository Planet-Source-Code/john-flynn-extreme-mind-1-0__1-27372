VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtremeMind 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extreme Mind"
   ClientHeight    =   7830
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   2880
   Icon            =   "Extreme Mind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Extreme Mind.frx":0C42
   ScaleHeight     =   7830
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3720
      Top             =   3600
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   7545
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1720
            MinWidth        =   1720
            Object.ToolTipText     =   "Current Difficulty Setting"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Object.ToolTipText     =   "Current Row Number"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Elapsed Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraWin 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   55
      TabIndex        =   2
      Top             =   55
      Width           =   2775
      Begin VB.Image imgWinPeg 
         Height          =   480
         Index           =   0
         Left            =   600
         Picture         =   "Extreme Mind.frx":2B27
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgWinPeg 
         Height          =   480
         Index           =   1
         Left            =   1175
         Picture         =   "Extreme Mind.frx":2EB7
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgWinPeg 
         Height          =   480
         Index           =   2
         Left            =   1750
         Picture         =   "Extreme Mind.frx":3247
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgWinPeg 
         Height          =   480
         Index           =   3
         Left            =   2300
         Picture         =   "Extreme Mind.frx":35D7
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheckCode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Check Code"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Image SelectedPeg 
      Height          =   480
      Left            =   3720
      Picture         =   "Extreme Mind.frx":3967
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image imgWhiteStatus 
      Height          =   165
      Left            =   4200
      Picture         =   "Extreme Mind.frx":3CF7
      Top             =   6840
      Width           =   165
   End
   Begin VB.Image imgBlackStatus 
      Height          =   165
      Left            =   3960
      Picture         =   "Extreme Mind.frx":3D6A
      Top             =   6840
      Width           =   165
   End
   Begin VB.Image imgBlankStatus 
      Height          =   165
      Left            =   3720
      Picture         =   "Extreme Mind.frx":3DDD
      Top             =   6840
      Width           =   165
   End
   Begin VB.Image HomePeg 
      DragIcon        =   "Extreme Mind.frx":412D
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   6
      Left            =   2375
      Picture         =   "Extreme Mind.frx":49F7
      Top             =   6325
      Width           =   480
   End
   Begin VB.Image HomePeg 
      DragIcon        =   "Extreme Mind.frx":4D92
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   5
      Left            =   1900
      Picture         =   "Extreme Mind.frx":565C
      Top             =   6325
      Width           =   480
   End
   Begin VB.Image HomePeg 
      DragIcon        =   "Extreme Mind.frx":59F7
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   4
      Left            =   1425
      Picture         =   "Extreme Mind.frx":62C1
      Top             =   6325
      Width           =   480
   End
   Begin VB.Image HomePeg 
      DragIcon        =   "Extreme Mind.frx":665C
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   3
      Left            =   950
      Picture         =   "Extreme Mind.frx":6F26
      Top             =   6325
      Width           =   480
   End
   Begin VB.Image HomePeg 
      DragIcon        =   "Extreme Mind.frx":72C1
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   2
      Left            =   475
      Picture         =   "Extreme Mind.frx":7B8B
      Top             =   6325
      Width           =   480
   End
   Begin VB.Image HomePeg 
      DragIcon        =   "Extreme Mind.frx":7F26
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   1
      Left            =   25
      Picture         =   "Extreme Mind.frx":87F0
      Top             =   6325
      Width           =   480
   End
   Begin VB.Image HomePeg 
      Height          =   480
      Index           =   0
      Left            =   2880
      Picture         =   "Extreme Mind.frx":8B8B
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   39
      Left            =   350
      Picture         =   "Extreme Mind.frx":8F1B
      Top             =   705
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   38
      Left            =   120
      Picture         =   "Extreme Mind.frx":926B
      Top             =   705
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   37
      Left            =   120
      Picture         =   "Extreme Mind.frx":95BB
      Top             =   900
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   36
      Left            =   350
      Picture         =   "Extreme Mind.frx":990B
      Top             =   900
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   35
      Left            =   350
      Picture         =   "Extreme Mind.frx":9C5B
      Top             =   1230
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   34
      Left            =   120
      Picture         =   "Extreme Mind.frx":9FAB
      Top             =   1230
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   33
      Left            =   120
      Picture         =   "Extreme Mind.frx":A2FB
      Top             =   1440
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   32
      Left            =   350
      Picture         =   "Extreme Mind.frx":A64B
      Top             =   1440
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   31
      Left            =   350
      Picture         =   "Extreme Mind.frx":A99B
      Top             =   1800
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   30
      Left            =   120
      Picture         =   "Extreme Mind.frx":ACEB
      Top             =   1800
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   29
      Left            =   120
      Picture         =   "Extreme Mind.frx":B03B
      Top             =   2040
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   28
      Left            =   350
      Picture         =   "Extreme Mind.frx":B38B
      Top             =   2040
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   27
      Left            =   350
      Picture         =   "Extreme Mind.frx":B6DB
      Top             =   2380
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   26
      Left            =   120
      Picture         =   "Extreme Mind.frx":BA2B
      Top             =   2380
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   25
      Left            =   120
      Picture         =   "Extreme Mind.frx":BD7B
      Top             =   2600
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   24
      Left            =   350
      Picture         =   "Extreme Mind.frx":C0CB
      Top             =   2600
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   23
      Left            =   350
      Picture         =   "Extreme Mind.frx":C41B
      Top             =   2950
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   22
      Left            =   120
      Picture         =   "Extreme Mind.frx":C76B
      Top             =   2950
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   21
      Left            =   120
      Picture         =   "Extreme Mind.frx":CABB
      Top             =   3175
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   20
      Left            =   350
      Picture         =   "Extreme Mind.frx":CE0B
      Top             =   3175
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   19
      Left            =   350
      Picture         =   "Extreme Mind.frx":D15B
      Top             =   3510
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   18
      Left            =   120
      Picture         =   "Extreme Mind.frx":D4AB
      Top             =   3510
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   17
      Left            =   120
      Picture         =   "Extreme Mind.frx":D7FB
      Top             =   3720
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   16
      Left            =   350
      Picture         =   "Extreme Mind.frx":DB4B
      Top             =   3720
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   15
      Left            =   350
      Picture         =   "Extreme Mind.frx":DE9B
      Top             =   4080
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   14
      Left            =   120
      Picture         =   "Extreme Mind.frx":E1EB
      Top             =   4080
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   13
      Left            =   120
      Picture         =   "Extreme Mind.frx":E53B
      Top             =   4320
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   12
      Left            =   350
      Picture         =   "Extreme Mind.frx":E88B
      Top             =   4320
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   11
      Left            =   350
      Picture         =   "Extreme Mind.frx":EBDB
      Top             =   4650
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   10
      Left            =   120
      Picture         =   "Extreme Mind.frx":EF2B
      Top             =   4650
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   9
      Left            =   120
      Picture         =   "Extreme Mind.frx":F27B
      Top             =   4875
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   8
      Left            =   350
      Picture         =   "Extreme Mind.frx":F5CB
      Top             =   4875
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   7
      Left            =   350
      Picture         =   "Extreme Mind.frx":F91B
      Top             =   5225
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   6
      Left            =   120
      Picture         =   "Extreme Mind.frx":FC6B
      Top             =   5225
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   5
      Left            =   120
      Picture         =   "Extreme Mind.frx":FFBB
      Top             =   5450
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   4
      Left            =   350
      Picture         =   "Extreme Mind.frx":1030B
      Top             =   5450
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   3
      Left            =   350
      Picture         =   "Extreme Mind.frx":1065B
      Top             =   5805
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   2
      Left            =   120
      Picture         =   "Extreme Mind.frx":109AB
      Top             =   5800
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   1
      Left            =   120
      Picture         =   "Extreme Mind.frx":10CFB
      Top             =   6015
      Width           =   165
   End
   Begin VB.Image Status 
      Height          =   165
      Index           =   0
      Left            =   350
      Picture         =   "Extreme Mind.frx":1104B
      Top             =   6015
      Width           =   165
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   39
      Left            =   2350
      Picture         =   "Extreme Mind.frx":1139B
      Top             =   625
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   38
      Left            =   1775
      Picture         =   "Extreme Mind.frx":1172B
      Top             =   625
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   37
      Left            =   1200
      Picture         =   "Extreme Mind.frx":11ABB
      Top             =   625
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   36
      Left            =   650
      Picture         =   "Extreme Mind.frx":11E4B
      Top             =   625
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   35
      Left            =   2350
      Picture         =   "Extreme Mind.frx":121DB
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   34
      Left            =   1775
      Picture         =   "Extreme Mind.frx":1256B
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   33
      Left            =   1200
      Picture         =   "Extreme Mind.frx":128FB
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   32
      Left            =   650
      Picture         =   "Extreme Mind.frx":12C8B
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   31
      Left            =   2350
      Picture         =   "Extreme Mind.frx":1301B
      Top             =   1750
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   30
      Left            =   1775
      Picture         =   "Extreme Mind.frx":133AB
      Top             =   1750
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   29
      Left            =   1200
      Picture         =   "Extreme Mind.frx":1373B
      Top             =   1750
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   28
      Left            =   650
      Picture         =   "Extreme Mind.frx":13ACB
      Top             =   1750
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   27
      Left            =   2350
      Picture         =   "Extreme Mind.frx":13E5B
      Top             =   2325
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   26
      Left            =   1775
      Picture         =   "Extreme Mind.frx":141EB
      Top             =   2325
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   25
      Left            =   1200
      Picture         =   "Extreme Mind.frx":1457B
      Top             =   2325
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   24
      Left            =   650
      Picture         =   "Extreme Mind.frx":1490B
      Top             =   2325
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   23
      Left            =   2350
      Picture         =   "Extreme Mind.frx":14C9B
      Top             =   2900
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   22
      Left            =   1775
      Picture         =   "Extreme Mind.frx":1502B
      Top             =   2900
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   21
      Left            =   1200
      Picture         =   "Extreme Mind.frx":153BB
      Top             =   2900
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   20
      Left            =   650
      Picture         =   "Extreme Mind.frx":1574B
      Top             =   2900
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   19
      Left            =   2350
      Picture         =   "Extreme Mind.frx":15ADB
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   18
      Left            =   1775
      Picture         =   "Extreme Mind.frx":15E6B
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   17
      Left            =   1200
      Picture         =   "Extreme Mind.frx":161FB
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   16
      Left            =   650
      Picture         =   "Extreme Mind.frx":1658B
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   15
      Left            =   2350
      Picture         =   "Extreme Mind.frx":1691B
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   14
      Left            =   1775
      Picture         =   "Extreme Mind.frx":16CAB
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   13
      Left            =   1200
      Picture         =   "Extreme Mind.frx":1703B
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   12
      Left            =   650
      Picture         =   "Extreme Mind.frx":173CB
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   11
      Left            =   2350
      Picture         =   "Extreme Mind.frx":1775B
      Top             =   4600
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   10
      Left            =   1775
      Picture         =   "Extreme Mind.frx":17AEB
      Top             =   4600
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   9
      Left            =   1200
      Picture         =   "Extreme Mind.frx":17E7B
      Top             =   4600
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   8
      Left            =   650
      Picture         =   "Extreme Mind.frx":1820B
      Top             =   4600
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   7
      Left            =   2350
      Picture         =   "Extreme Mind.frx":1859B
      Top             =   5200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   6
      Left            =   1775
      Picture         =   "Extreme Mind.frx":1892B
      Top             =   5200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   5
      Left            =   1200
      Picture         =   "Extreme Mind.frx":18CBB
      Top             =   5200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   4
      Left            =   650
      Picture         =   "Extreme Mind.frx":1904B
      Top             =   5200
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   3
      Left            =   2355
      Picture         =   "Extreme Mind.frx":193DB
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   2
      Left            =   1770
      Picture         =   "Extreme Mind.frx":1976B
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "Extreme Mind.frx":19AFB
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Peg 
      Height          =   480
      Index           =   0
      Left            =   650
      Picture         =   "Extreme Mind.frx":19E8B
      Top             =   5760
      Width           =   480
   End
   Begin VB.Menu Game 
      Caption         =   "&Game"
      Begin VB.Menu new 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu difficulty 
         Caption         =   "&Difficulty"
         Begin VB.Menu easy 
            Caption         =   "&Easy"
            Shortcut        =   ^E
         End
         Begin VB.Menu hard 
            Caption         =   "&Hard"
            Shortcut        =   ^H
         End
         Begin VB.Menu extreme 
            Caption         =   "Extre&me"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu sound 
         Caption         =   "&Sound"
         Begin VB.Menu sndOn 
            Caption         =   "On"
         End
         Begin VB.Menu sndOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu cheat 
         Caption         =   "&Cheat"
      End
      Begin VB.Menu spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu scores 
         Caption         =   "&View High Scores"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu contents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmExtremeMind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this is the state of the web browser 2 stands for normal and 3 would be maximized
Private Const SW_SHOWNORMAL = 2

'declare the function to launch a web browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub cheat_Click()
    frmCheat.Show
End Sub

Private Sub cmdCheckCode_Click()
    Dim i, c, LastPeg, black, X, Y As Integer
    Dim skip(0 To 3) As Boolean
    black = 0
    LastPeg = (ActiveRow * 4)
    
    'only check the code if the answers have not already been shown
    If fraWin.Visible = False Then
    
    ' this if statement will look for errors associated with
    ' the current difficulty setting
    If easy.Checked = True Then
        If Peg(LastPeg - 4).Picture = HomePeg(0).Picture Or Peg(LastPeg - 3).Picture = HomePeg(0).Picture Or Peg(LastPeg - 2).Picture = HomePeg(0).Picture Or Peg(LastPeg - 1).Picture = HomePeg(0).Picture Then
            MsgBox "No empty pegs are allowed on this difficulty setting.", vbExclamation, "You have 1 or more empty pegs"
            Exit Sub
        End If
        
        'check for duplicates
        For i = (LastPeg - 4) To (LastPeg - 1)
            For X = i + 1 To (LastPeg - 1)
                If Peg(X) = Peg(i) Then
                    MsgBox "No duplicate pegs are allowed on this difficulty setting", vbExclamation, "You have 1 or more duplicate pegs"
                    Exit Sub
                End If
            Next
        Next
        
    ElseIf hard.Checked = True Then
        If Peg(LastPeg - 4).Picture = HomePeg(0).Picture Or Peg(LastPeg - 3).Picture = HomePeg(0).Picture Or Peg(LastPeg - 2).Picture = HomePeg(0).Picture Or Peg(LastPeg - 1).Picture = HomePeg(0).Picture Then
            MsgBox "No empty pegs are allowed on this difficulty setting.", vbExclamation, "You have 1 or more empty pegs"
            Exit Sub
        End If
    End If
    
    'initalize the skip array to false
    For X = 0 To 3
        skip(X) = False
    Next
    
    'first turn on any white status pegs
    X = 0
    i = LastPeg - 4
    Y = i
    Do
        c = LastPeg - 4
        Do
            If Peg(c).Picture = imgWinPeg(X).Picture And skip(X) = False And flag(c) = False Then
                Status(Y) = imgWhiteStatus.Picture
                Y = Y + 1
                flag(c) = True 'flag this spot so it doesn't get checked again
                skip(X) = True 'flag this spot so it doesn't get checked again
            End If
        c = c + 1
        Loop While c < LastPeg
    X = X + 1
    i = i + 1
    Loop While i < LastPeg
    
    'now turn on any black status pegs
    i = LastPeg - 4
    Y = i
    For c = 0 To 3
        If Peg(i).Picture = imgWinPeg(c).Picture Then
            Status(Y).Picture = imgBlackStatus.Picture
            black = black + 1
            Y = Y + 1
        End If
        i = i + 1
    Next
    
    'see if the user got the correct answer
    If black = 4 Then
        Timer.Enabled = False
        fraWin.Visible = True
        If CurrentCheat = True Then
            frmCheatWin.Show
        Else
            CurrentScore.Row = ActiveRow
            CurrentScore.Time = ElapsedTime
            If easy.Checked = True Then
                CurrentScore.Level = 3
            ElseIf hard.Checked = True Then
                CurrentScore.Level = 2
            ElseIf extreme.Checked = True Then
                CurrentScore.Level = 1
            End If
            frmWin.Show
        End If
        Exit Sub
    End If
    
    'get the next row ready to go
    Call NextRow
    
    End If 'this is where it will exit if the answers are already shown
End Sub

Private Sub cmdExit_Click()
    Unload frmExtremeMind
End Sub

Private Sub contents_Click()
    'Launch the default web browser and navigate to the help file
    '*NOTE* you could replace the App.path & "\Help\index.html" with a simple "http://www.website.com"
    ShellExecute Me.hWnd, "open", App.Path & "\Help\index.html", ByVal 0&, "", SW_SHOWNORMAL
End Sub

Private Sub easy_Click()
Dim n As Integer
    
    'un-check hard, check easy and start a new game
    If hard.Checked = True Then
        n = MsgBox("This will start a new game. Are you sure you want to abort the current game?", vbYesNo, "New Game?")
        If n = 7 Then '7 is returned when the user clicks no
            Exit Sub
        ElseIf n = 6 Then '6 is returned when the user clicks yes
            hard.Checked = False
            easy.Checked = True
            Call NewGame
            'update the status bar message
            stat.Panels(1).Text = "Easy"
        End If
    
    'un-check extreme, check easy and start a new game
    ElseIf extreme.Checked = True Then
        n = MsgBox("This will start a new game. Are you sure you want to abort the current game?", vbYesNo, "New Game?")
        If n = 7 Then '7 is returned when the user clicks no
            Exit Sub
        ElseIf n = 6 Then '6 is returned when the user clicks yes
            extreme.Checked = False
            easy.Checked = True
            Call NewGame
            'update the status bar message
            stat.Panels(1).Text = "Easy"
        End If
    End If
    
    'Log this new setting so it will be remembered on exit
    Get #1, 11, CurrentScore
    CurrentScore.Level = 3
    Put #1, 11, CurrentScore
    
    
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'get the settings for the level and sound and then
    'activate the settings
    Get #1, 11, CurrentScore
    
    'activate the level
    If CurrentScore.Level = 3 Then
        easy.Checked = True
    ElseIf CurrentScore.Level = 2 Then
        hard.Checked = True
    ElseIf CurrentScore.Level = 1 Then
        extreme.Checked = True
    End If
    
    'activate the sound
    If CurrentScore.Row = 1 Then
        sndOn.Checked = True
        snd = True
    ElseIf CurrentScore.Row = 0 Then
        sndOff.Checked = True
        snd = False
    End If
        
    
    Call NewGame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'only allow the form to be dragged with the left mouse button
    If Button = 1 Then
        FormDrag Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'close the high score file before exiting the program
    Close #1
End Sub

Private Sub hard_Click()
Dim n As Integer
    
    'un-check easy, check hard and start a new game
    If easy.Checked = True Then
        n = MsgBox("This will start a new game. Are you sure you want to abort the current game?", vbYesNo, "New Game?")
        If n = 7 Then '7 is returned when the user clicks no
            Exit Sub
        ElseIf n = 6 Then '6 is returned when the user clicks yes
            easy.Checked = False
            hard.Checked = True
            Call NewGame
            'update the status bar message
            stat.Panels(1).Text = "Hard"
        End If
    
    'un-check extreme, check hard and start a new game
    ElseIf extreme.Checked = True Then
        n = MsgBox("This will start a new game. Are you sure you want to abort the current game?", vbYesNo, "New Game?")
        If n = 7 Then '7 is returned when the user clicks no
            Exit Sub
        ElseIf n = 6 Then '6 is returned when the user clicks yes
            extreme.Checked = False
            hard.Checked = True
            Call NewGame
            'update the status bar message
            stat.Panels(1).Text = "Hard"
        End If
    End If
    
    'Log this new setting so it will be remembered on exit
    Get #1, 11, CurrentScore
    CurrentScore.Level = 2
    Put #1, 11, CurrentScore

End Sub

Private Sub extreme_Click()
Dim n As Integer
    
    'un-check easy, check extreme and start a new game
    If easy.Checked = True Then
        n = MsgBox("This will start a new game. Are you sure you want to abort the current game?", vbYesNo, "New Game?")
        If n = 7 Then '7 is returned when the user clicks no
            Exit Sub
        ElseIf n = 6 Then '6 is returned when the user clicks yes
            easy.Checked = False
            extreme.Checked = True
            Call NewGame
            'update the status bar message
            stat.Panels(1).Text = "Extreme"
        End If
    
    'un-check hard, check extreme and start a new game
    ElseIf hard.Checked = True Then
        n = MsgBox("This will start a new game. Are you sure you want to abort the current game?", vbYesNo, "New Game?")
        If n = 7 Then '7 is returned when the user clicks no
            Exit Sub
        ElseIf n = 6 Then '6 is returned when the user clicks yes
            hard.Checked = False
            extreme.Checked = True
            Call NewGame
            'update the status bar message
            stat.Panels(1).Text = "Extreme"
        End If
    End If
    
    'Log this new setting so it will be remembered on exit
    Get #1, 11, CurrentScore
    CurrentScore.Level = 1
    Put #1, 11, CurrentScore

End Sub

Private Sub new_Click()
    Call NewGame
End Sub

Private Sub NewGame()
    Call GetCode
    
    'set the active row to row 0 so the call of NextRow will work
    ActiveRow = 0
    
    'set the cheat to false
    CurrentCheat = False
    
    'enable the next row
    Call NextRow
    
    'place the current difficulty setting in the first panel of the status bar
    If easy.Checked = True Then
        stat.Panels(1).Text = "Easy"
    ElseIf hard.Checked = True Then
        stat.Panels(1).Text = "Hard"
    ElseIf extreme.Checked = True Then
        stat.Panels(1).Text = "Extreme"
    End If
    
    'enable the timer and set it to zero
    ElapsedTime = 0
    stat.Panels(3).Text = ElapsedTime
    Timer.Enabled = True
    
End Sub

Private Sub GetCode()
    Randomize
    
    Dim i As Integer
    Dim color As Integer
    
    'hide the answers
    fraWin.Visible = False
    
    'get random numbers with no duplicates and no zeros
    If easy.Checked = True Then
        Dim counter, c, hold, numbers(3) As Integer
        
        counter = 0
        c = 0
        hold = 0
        
        'zero out the array
        For i = 0 To 3
            numbers(i) = 0
        Next
        
        Do 'start the algorithm for getting numbers without duplicates
            hold = Int(6 * Rnd + 1)
            i = 0
            Do
                If hold = numbers(i) Then 'duplicate found
                    Exit Do
                ElseIf i = c Then 'got to the next available spot in the array
                    numbers(c) = hold
                    hold = 0 'because zero is not included in the random numbers
                    c = c + 1 'increase c only when no duplicates were found
                    Exit Do
                End If
                i = i + 1 'increase i
            Loop While i <= c
            If hold = 0 Then
                counter = counter + 1 'increase counter because next position was filled
            End If
        Loop While counter <= 3 'ending of the algorithm for no duplicates
        
        'place the images in the right spots
        For i = 0 To 3
            imgWinPeg(i).Picture = HomePeg(numbers(i)).Picture
        Next
        
    'get random numbers that can re-peat but don't include zeros
    ElseIf hard.Checked = True Then
        For i = 0 To 3
            'Get a random color value
            color = Int(6 * Rnd + 1)
            'Place the color corresponding to the random number
            imgWinPeg(i).Picture = HomePeg(color).Picture
        Next
        
    'get random numbers that can re-peat and include zeros
    ElseIf extreme.Checked = True Then
        For i = 0 To 3
            'Get a random color value
            color = Int(6 * Rnd + 0)
            'Place the color corresponding to the random number
            imgWinPeg(i).Picture = HomePeg(color).Picture
        Next
    End If
    
End Sub

Private Sub NextRow()
    Dim i, LastPeg As Integer
    LastPeg = 39
    
    'the user lost the game
    If ActiveRow = 10 Then
        'show the answers and disable the timer
        Timer.Enabled = False
        fraWin.Visible = True
        frmLose.Show
        Exit Sub
    
    'enable the first row and disable all others
    ElseIf ActiveRow = 0 Then
        For i = 0 To 3
            Peg(i).Enabled = True
        Next
        
        For i = 4 To LastPeg
            Peg(i).Enabled = False
        Next
        
        'reset all the images and the flag array
        For i = 0 To LastPeg
            flag(i) = False
            Peg(i).Picture = HomePeg(0).Picture
            Status(i).Picture = imgBlankStatus.Picture
        Next
        
        ActiveRow = ActiveRow + 1
    
    'enable next row and disable previous row
    Else
        i = (ActiveRow * 4)
        LastPeg = i - 4
        Do
            Peg(i).Enabled = False
            i = i - 1
        Loop While i >= LastPeg
        
        i = (ActiveRow * 4)
        LastPeg = i + 4
        Do
            Peg(i).Enabled = True
            i = i + 1
        Loop While i < LastPeg
        
        ActiveRow = ActiveRow + 1
    End If
    
    stat.Panels(2).Text = "Row " & ActiveRow
    
End Sub

Private Sub Peg_DblClick(Index As Integer)
    If Peg(Index).Picture <> HomePeg(0).Picture Then
        Peg(Index).Picture = HomePeg(0).Picture
    End If
End Sub

Private Sub Peg_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    
    'display the color peg that was dropped
    Peg(Index).Picture = Source.Picture
    SelectedPeg.Picture = Peg(Index).Picture
    
    If snd = True Then
        'play a sound
        PlaySound App.Path & "\Sounds\drop.wav"
    End If
    
End Sub

Private Sub Peg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Peg(Index).Picture = HomePeg(0).Picture Then
            Peg(Index).Picture = SelectedPeg.Picture
            If snd = True Then
                'play a sound
                PlaySound App.Path & "\Sounds\drop.wav"
            End If
        Else
            SelectedPeg.Picture = Peg(Index).Picture
        End If
        
    ElseIf Button = 2 Then
        If Peg(Index).Picture <> HomePeg(0).Picture Then
            SelectedPeg.Picture = Peg(Index).Picture
            Peg(Index).Picture = HomePeg(0).Picture
        End If
    End If
End Sub

Private Sub scores_Click()
    frmHighScores.Show
End Sub

Private Sub sndoff_Click()
    If sndOn.Checked = True Then
        sndOn.Checked = False
        sndOff.Checked = True
        snd = False
    End If
    
    'Log this new setting so it will be remembered on exit
    Get #1, 11, CurrentScore
    CurrentScore.Row = 0
    Put #1, 11, CurrentScore
End Sub

Private Sub sndon_Click()
    If sndOff.Checked = True Then
        sndOff.Checked = False
        sndOn.Checked = True
        snd = True
    End If
    
    'Log this new setting so it will be remembered on exit
    Get #1, 11, CurrentScore
    CurrentScore.Row = 1
    Put #1, 11, CurrentScore
End Sub

Private Sub Timer_Timer()
    ElapsedTime = ElapsedTime + 1
    stat.Panels(3).Text = ElapsedTime
End Sub
