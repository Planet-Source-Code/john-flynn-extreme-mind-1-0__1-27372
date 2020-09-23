VERSION 5.00
Begin VB.Form frmWin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Good Job!"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "Win.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O&k"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
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
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label lblWin 
      BackStyle       =   0  'Transparent
      Caption         =   "Good job you cracked the code."
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
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image imgWin 
      Height          =   480
      Index           =   3
      Left            =   2400
      Picture         =   "Win.frx":0C42
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWin 
      Height          =   480
      Index           =   2
      Left            =   1800
      Picture         =   "Win.frx":0FD2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWin 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "Win.frx":1362
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWin 
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "Win.frx":16F2
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "frmWin"
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
        imgWin(i).Picture = frmExtremeMind.imgWinPeg(i).Picture
    Next
    
    If snd = True Then
        'play a sound
        PlaySound App.Path & "\Sounds\Jackpot.wav"
    End If
    
    Call ScanScores
    
End Sub
