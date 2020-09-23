VERSION 5.00
Begin VB.Form frmCheat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheater"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "Cheat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O&k"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCheat 
      Height          =   480
      Index           =   3
      Left            =   2400
      Picture         =   "Cheat.frx":0C42
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgCheat 
      Height          =   480
      Index           =   2
      Left            =   1800
      Picture         =   "Cheat.frx":0FD2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgCheat 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "Cheat.frx":1362
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgCheat 
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "Cheat.frx":16F2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The code is:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheaters never win, and winners never cheat!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmCheat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    CurrentCheat = True
    'Load the winning pictures
    For i = 0 To 3
        imgCheat(i).Picture = frmExtremeMind.imgWinPeg(i).Picture
    Next

End Sub
