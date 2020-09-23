VERSION 5.00
Begin VB.Form frmCheatWin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You Cheated"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CheatWin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O&k"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1320
      Picture         =   "CheatWin.frx":0C42
      Top             =   675
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The code was:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Good job, but next time see if you can do it without cheating"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   2520
      Picture         =   "CheatWin.frx":0FBD
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "CheatWin.frx":134D
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1320
      Picture         =   "CheatWin.frx":16DD
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   720
      Picture         =   "CheatWin.frx":1A6D
      Top             =   1560
      Width           =   480
   End
End
Attribute VB_Name = "frmCheatWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'Load the winning pictures
    For i = 0 To 3
        Image1(i).Picture = frmExtremeMind.imgWinPeg(i).Picture
    Next
    
    If snd = True Then
        'play a sound
        PlaySound App.Path & "\Sounds\Glass.wav"
    End If
End Sub
