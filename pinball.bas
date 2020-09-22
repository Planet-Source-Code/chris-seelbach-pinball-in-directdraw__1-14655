Attribute VB_Name = "Module1"
Const vbKeyEscape = &H1B
Const vbKeySpace = &H20
Const vbKeyF8 = &H77
Const vbKeyA = 65
Const vbKeyL = 76
'sound calls
Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
'the authors score (my eye's are not
'what they used to be)
Global MyScore As Long
'the all time high score
Global HighScore As Long
'bonus ball variable
Global n As Integer
'the ball #
Global BallNumber As Integer
'load the next ball
Global LoadBall As Integer
'your score
Global Score As Long
'if were finished
Global Quit As Integer
'credit screen declares
Global CreditA As Long
Global CreditB As Long
'plunger strength
Global Plunger As Long
'detain the ball
Global Detain As Integer
'bonus ring
Global Ring As Integer
'pull the plunger
Global Pull As Long
'shot the ball and the game is started
Global Shoot As Long
'when you begin playing
Global Begin As Long
'you lost!
Global Lost As Integer
'the ball is committed
Global Commit As Integer
Global Inplay As Integer
'if you pushed a flipper button
Global Rf As Integer
Global Lf As Integer
Global Flipper As Integer
'velocity & drag constants
Global v As Long
Global d As Long
Global g As Long
Global b As Long
Global h1 As Long
Global v1 As Long
'random integer
Global r As Integer
Global pb As Integer




