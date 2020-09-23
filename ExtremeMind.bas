Attribute VB_Name = "modExtremeMind"
Option Explicit

'some useful color constants that are a slightly different shade than the
'default VB color constants
Public Const BLUE = &HFF8080
Public Const RED = &H8080FF
Public Const GREEN = &H80FF80
Public Const YELLOW = &H80FFFF

'these 2 were not used on this project
Public Const PURPLE = &HFF80FF
Public Const TAN = &HC0E0FF

' declare the global variables
Public ActiveRow As Integer
Public CurrentCheat As Boolean
Public snd As Boolean
Public ElapsedTime As Integer
Public flag(0 To 39) As Boolean

'these 2 will allow a form to be drag and dropped by dragging anywhere on one of the form
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()

'define the udt(user defined type) for the high scores
Type HighScores
    Name As String * 25
    Level As Integer
    Row As Integer
    Time As Integer
End Type

'CheckScore will be used for comparing the current game
'to the existing scores and CurrentScore will contain the
'scores for the current user. HoldScore will be used when
'swapping the scores around in the HighScores.dat file
Public CheckScore As HighScores
Public HoldScore As HighScores
Public CurrentScore As HighScores
Public Rank As Integer
Public Path As String
Public UserName As String


'here is the start-up object
Public Sub Main()
    ActiveRow = 0
    snd = True
    
    Path = (App.Path & "\Extreme Mind.dat")
    
    'open the high score list and if it doesn't exist it will be created
    'and each record will be set to the length of one of the HighScores udt
    If Dir(Path) = "" Then
        Open (Path) For Random As #1 Len = Len(CheckScore)
        InitHighScores
    Else
        Open (Path) For Random As #1 Len = Len(CheckScore)
    End If
    
    frmExtremeMind.Show
End Sub

Public Sub FormDrag(frmFormToDrag As Form)
    ReleaseCapture
    Call SendMessage(frmFormToDrag.hwnd, &HA1, 2, 0&)
End Sub

Public Sub ScanScores()
    
    'the level is the most important so if the user is playing extreme and the highest score is on easy
    'the user will get the highest score regardless of the row or time. Likewise the row is more
    'important than the time and if the highest score has more rows than the user the user will slide
    'their score into the file before that one

    For Rank = 1 To 10
        Get #1, Rank, CheckScore
        If CurrentScore.Level < CheckScore.Level Then
            EnterScore
            Exit Sub
        ElseIf CurrentScore.Level = CheckScore.Level Then
            If CurrentScore.Row < CheckScore.Row Then
                EnterScore
                Exit Sub
            ElseIf CurrentScore.Row = CheckScore.Row Then
                If CurrentScore.Time < CheckScore.Time Then
                    EnterScore
                    Exit Sub
                End If
            End If
        End If
    Next
    
End Sub

Public Sub EnterScore()
    Dim i As Integer
    
    'get the users name
    CurrentScore.Name = InputBox("You made the high score list. Please enter your name.", "Good Job!")
    
    'get the selected score from the file before entering the user's score
    'Get #1, Rank, CheckScore
    
    'place the user's record in the current spot
    Put #1, Rank, CurrentScore
                    
    'move to the next record
    Rank = Rank + 1
    
    'move the other records down on the list
    For i = Rank To 10
        Get #1, i, HoldScore
        Put #1, i, CheckScore
        CheckScore = HoldScore
    Next
    
    frmHighScores.Show
End Sub

Public Sub InitHighScores()
    
    'initialize the high score file with data that will show on a new high score list
    
    CurrentScore.Name = "Empty"
    CurrentScore.Level = 3 'Easy
    CurrentScore.Row = 11
    CurrentScore.Time = 9999
    
    For Rank = 1 To 10
        Put #1, Rank, CurrentScore
    Next
    
    'Note: The eleventh entry will be used for remembering the user settings on exit
    'the only settings being remembered are the current level and sound. The level will
    'be the same as it is for the high score list and the row is used for the sound.
    'Row = 1 will be the default meaning sound is on and 0 will be used when the sound is
    'turned off
    CurrentScore.Row = 1
    Put #1, 11, CurrentScore
    
End Sub
