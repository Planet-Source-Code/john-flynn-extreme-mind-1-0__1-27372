VERSION 5.00
Begin VB.Form frmLose 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sorry!"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   Icon            =   "Lose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O&k"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The code was:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1425
   End
   Begin VB.Image imgLose 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "Lose.frx":0C42
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgLose 
      Height          =   480
      Index           =   2
      Left            =   1680
      Picture         =   "Lose.frx":0FD2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgLose 
      Height          =   480
      Index           =   1
      Left            =   1080
      Picture         =   "Lose.frx":1362
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgLose 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "Lose.frx":16F2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblLose 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorry, you didn't get the code!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLose"
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
        imgLose(i).Picture = frmExtremeMind.imgWinPeg(i).Picture
    Next
    
    If snd = True Then
        'play a sound
        PlaySound App.Path & "\Sounds\Glass.wav"
    End If
    
End Sub
