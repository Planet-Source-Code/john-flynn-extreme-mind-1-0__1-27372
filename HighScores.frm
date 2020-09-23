VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "HighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Line Line15 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   240
      TabIndex        =   54
      Top             =   2400
      Width           =   240
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   240
      TabIndex        =   53
      Top             =   2160
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   240
      TabIndex        =   52
      Top             =   1920
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   51
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   50
      Top             =   1440
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   49
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   48
      Top             =   960
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   47
      Top             =   720
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   46
      Top             =   480
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   45
      Top             =   240
      Width           =   120
   End
   Begin VB.Label lblRank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rank"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   465
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   720
      X2              =   720
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Line Line5 
      X1              =   8880
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   4920
      TabIndex        =   43
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   4920
      TabIndex        =   42
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   4920
      TabIndex        =   41
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   4920
      TabIndex        =   40
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   4920
      TabIndex        =   39
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4920
      TabIndex        =   38
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4920
      TabIndex        =   37
      Top             =   960
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4920
      TabIndex        =   36
      Top             =   720
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4920
      TabIndex        =   35
      Top             =   480
      Width           =   465
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4920
      TabIndex        =   34
      Top             =   240
      Width           =   465
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   4080
      TabIndex        =   33
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   4080
      TabIndex        =   32
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   4080
      TabIndex        =   31
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   4080
      TabIndex        =   30
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   4080
      TabIndex        =   29
      Top             =   1440
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4080
      TabIndex        =   28
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4080
      TabIndex        =   27
      Top             =   960
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4080
      TabIndex        =   26
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4080
      TabIndex        =   25
      Top             =   480
      Width           =   405
   End
   Begin VB.Label lblRow 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4080
      TabIndex        =   24
      Top             =   240
      Width           =   405
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   2760
      TabIndex        =   23
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   2760
      TabIndex        =   22
      Top             =   2160
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2760
      TabIndex        =   21
      Top             =   1920
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2760
      TabIndex        =   20
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   2760
      TabIndex        =   19
      Top             =   1440
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2760
      TabIndex        =   17
      Top             =   960
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2760
      TabIndex        =   16
      Top             =   720
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   480
      Width           =   870
   End
   Begin VB.Label lblDifficulty 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2760
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   960
      TabIndex        =   13
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   960
      TabIndex        =   12
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   960
      TabIndex        =   11
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   960
      TabIndex        =   10
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   960
      TabIndex        =   9
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   8880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   0
      Width           =   1860
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   465
   End
   Begin VB.Label lblRow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   405
   End
   Begin VB.Label lblDifficulty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   240
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To 10
        Get #1, i, CheckScore
        lblName(i).Caption = CheckScore.Name
        If CheckScore.Level = 1 Then
            lblDifficulty(i).Caption = "Extreme"
        ElseIf CheckScore.Level = 2 Then
            lblDifficulty(i).Caption = "Hard"
        Else 'CheckScore.Level = 3 Then
            lblDifficulty(i).Caption = "Easy"
        End If
        lblRow(i).Caption = CheckScore.Row
        lblTime(i).Caption = CheckScore.Time
        
        'now change the font color to seperate the levels
        If lblDifficulty(i).Caption = "Extreme" Then
            lblRank(i).ForeColor = RED
            lblName(i).ForeColor = RED
            lblDifficulty(i).ForeColor = RED
            lblRow(i).ForeColor = RED
            lblTime(i).ForeColor = RED
        ElseIf lblDifficulty(i).Caption = "Hard" Then
            lblRank(i).ForeColor = BLUE
            lblName(i).ForeColor = BLUE
            lblDifficulty(i).ForeColor = BLUE
            lblRow(i).ForeColor = BLUE
            lblTime(i).ForeColor = BLUE
        ElseIf lblDifficulty(i).Caption = "Easy" Then
            lblRank(i).ForeColor = GREEN
            lblName(i).ForeColor = GREEN
            lblDifficulty(i).ForeColor = GREEN
            lblRow(i).ForeColor = GREEN
            lblTime(i).ForeColor = GREEN
        End If
        
        'now if it is the default value then don't show any information in the labels
        'and change the rank forecolor to YELLOW
        If lblTime(i).Caption = 9999 Then
            lblRank(i).ForeColor = YELLOW
            lblName(i).Caption = ""
            lblDifficulty(i).Caption = ""
            lblRow(i).Caption = ""
            lblTime(i).Caption = ""
        End If
    Next
End Sub
