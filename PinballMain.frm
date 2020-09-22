VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5625
   ClientLeft      =   2355
   ClientTop       =   1620
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "PinballMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "PinballMain.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2040
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************
'collision detection of the ball
'and the objects in the "machine"
'is designed almost entirely using
'the Point method. I wanted to see how
'effective it would be in a fast game.
'Hope you get a neat idea from the method.
'***********************************
Option Explicit
'DX stuff
Dim binit As Boolean
Dim dx As New DirectX7
Dim dd As DirectDraw7
Dim flagsurf As DirectDrawSurface7
Dim spritesurf As DirectDrawSurface7
Dim spritesurf1 As DirectDrawSurface7 'the ball
Dim spritesurf2 As DirectDrawSurface7 'the shooter
Dim spritesurf3 As DirectDrawSurface7 'future things to add
Dim spritesurf4 As DirectDrawSurface7 'future things to add
Dim spritesurf5 As DirectDrawSurface7 'trough cover
Dim spritesurf6 As DirectDrawSurface7 'right flipper
Dim spritesurf7 As DirectDrawSurface7 'left flipper
Dim spritesurf8 As DirectDrawSurface7 'table cover
Dim primary As DirectDrawSurface7
Dim backbuffer As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2
Dim ddsd4 As DDSURFACEDESC2
Dim ddsd5 As DDSURFACEDESC2
Dim ddsd6 As DDSURFACEDESC2
Dim ddsd7 As DDSURFACEDESC2
Dim ddsd8 As DDSURFACEDESC2
Dim ddsd9 As DDSURFACEDESC2
Dim ddsd10 As DDSURFACEDESC2
Dim ddsd11 As DDSURFACEDESC2
Dim ddsd12 As DDSURFACEDESC2 'right flipper
Dim ddsd13 As DDSURFACEDESC2 'left flipper
Dim ddsd14 As DDSURFACEDESC2
'sounds
Dim DSOUND As DirectSound
Dim BonusHit As DirectSoundBuffer
Dim LoseBall As DirectSoundBuffer
Dim BumperHit As DirectSoundBuffer
'objects
Dim spriteWidth As Integer
Dim spriteHeight As Integer
Dim cols As Integer
Dim rows As Integer
Dim row As Integer
Dim col As Integer
Dim RflipperFrame As Integer
Dim LflipperFrame As Integer
Dim brunning As Boolean
Dim CurModeActiveStatus As Boolean
Dim bRestore As Boolean







Sub Init()
    On Local Error GoTo errOut
    
    Dim file As String
    
    Set dd = dx.DirectDrawCreate("")
    Form1.Show
    
    'indicate that we dont need to change display depth
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    Call dd.SetDisplayMode(640, 480, 32, 0, DDSDM_DEFAULT)
    
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)
    
    'Get the backbuffer
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd4
         
    'Create DrawableSurface class form backbuffer
    'if you want to draw text and have the
    'background transparent;
    backbuffer.SetFontTransparency True
    'sound dec.
    Dim BufferDesc As DSBUFFERDESC
    Dim WavFormat As WAVEFORMATEX
    Set DSOUND = dx.DirectSoundCreate("")
    DSOUND.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    'load the sounds
    Set BonusHit = DSOUND.CreateSoundBufferFromFile(App.Path & "\ding.wav", BufferDesc, WavFormat)
    Set LoseBall = DSOUND.CreateSoundBufferFromFile(App.Path & "\oops.wav", BufferDesc, WavFormat)
    Set BumperHit = DSOUND.CreateSoundBufferFromFile(App.Path & "\bumper.wav", BufferDesc, WavFormat)

    'init the surfaces
    InitSurfaces
    
CreditA = 480 'the position of text on the credit screen
CreditB = 500

    binit = True
    brunning = True
    Do While brunning
        blt
        DoEvents
    Loop
errOut:
    EndIt
End Sub

Sub InitSurfaces()


    Set flagsurf = Nothing
    Set spritesurf = Nothing
    
   
    
    'load the bitmap into a surface -  the blank
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = ddsd4.lWidth
    ddsd2.lHeight = ddsd4.lHeight
    'we are starting the game:
    'to save file size, the intro screen was not
    'included so this is just a low bit blank
    Set flagsurf = dd.CreateSurfaceFromFile("blank.bmp", ddsd2)
    
    ddsd3.lFlags = DDSD_CAPS
    ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    Set spritesurf1 = dd.CreateSurfaceFromFile("ball.bmp", ddsd7)
    Set spritesurf2 = dd.CreateSurfaceFromFile("anglecover.bmp", ddsd8)
    Set spritesurf3 = dd.CreateSurfaceFromFile("shotcover.bmp", ddsd9)
    Set spritesurf4 = dd.CreateSurfaceFromFile("plunger.bmp", ddsd10)
    Set spritesurf5 = dd.CreateSurfaceFromFile("troughcover.bmp", ddsd11)
    Set spritesurf6 = dd.CreateSurfaceFromFile("rflipper.bmp", ddsd12) 'right flipper
    Set spritesurf7 = dd.CreateSurfaceFromFile("lflipper.bmp", ddsd13) 'left flipper
    Set spritesurf8 = dd.CreateSurfaceFromFile("tablecover.bmp", ddsd14)
    'the flippers
    spriteWidth = 33
    spriteHeight = 34
    cols = ddsd12.lWidth / spriteWidth 'ddsd3.lWidth / spriteWidth
    rows = ddsd12.lHeight / spriteHeight 'ddsd3.lHeight / spriteHeight
    
    'use black for transparent color key which is on
    'the source bitmap -> use src keying
    Dim key As DDCOLORKEY
    key.low = 0 'black
    key.high = 0 'black
    
    spritesurf1.SetColorKey DDCKEY_SRCBLT, key
    spritesurf2.SetColorKey DDCKEY_SRCBLT, key 'shotcover
    spritesurf3.SetColorKey DDCKEY_SRCBLT, key
    spritesurf4.SetColorKey DDCKEY_SRCBLT, key
    spritesurf5.SetColorKey DDCKEY_SRCBLT, key
    spritesurf6.SetColorKey DDCKEY_SRCBLT, key
    spritesurf7.SetColorKey DDCKEY_SRCBLT, key
    spritesurf8.SetColorKey DDCKEY_SRCBLT, key
    
End Sub


Sub blt()
    On Local Error GoTo errOut
    If binit = False Then Exit Sub
    
    Dim ddrval As Long
    Static i As Integer
    
    Dim rBack As RECT
    Dim rFlag As RECT
    Dim rSprite As RECT
    Dim rSprite2 As RECT
    Dim rSprite3 As RECT
    Dim rSprite4 As RECT
    Dim rSprite5 As RECT
    Dim rSprite6 As RECT
    Dim rSprite7 As RECT
    Dim rSprite8 As RECT 'right flipper
    Dim rSprite9 As RECT 'left flipper
    Dim rSprite10 As RECT 'table cover
    
    '
    Dim rPrim As RECT
    
    Static a As Single
    Static x As Single 'horizontal motion
    Static y As Single 'vertical motion
    Static g As Single 'gravity
    Static d As Single 'delay
    Static b As Single
    Static h1 As Single
    Static v1 As Single
    'timer
    Static t As Single
    Static t2 As Single
    Static tLast As Single
    'frame rate
    Static fps As Single
    'the ball your playing
    Static BallNumber As Integer
    'position some of the sprites initially
    rSprite6.Left = 428
    rSprite6.Top = 400
    rSprite3.Left = 406 'the angle cover
    rSprite3.Top = 196
    rSprite7.Left = 417 'the trough cover
    rSprite7.Top = 296
    rSprite8.Top = 390 'right flipper
    rSprite8.Left = 332
    rSprite9.Top = 388 'left flipper
    rSprite9.Left = 276
    rSprite10.Left = 266 'table cover
    rSprite10.Top = 423
    ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
    bRestore = False
    Do Until ExModeActive
        DoEvents
        bRestore = True
    Loop
    
    ' if we lost and got back the surfaces, then restore them
    DoEvents
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces ' must init the surfaces again if they we're lost
    End If
    
    'get the area of the screen where our window is
    rBack.Bottom = ddsd4.lHeight
    rBack.Right = ddsd4.lWidth
    
    'get the area of the bitmap we want ot blt
    rFlag.Bottom = ddsd2.lHeight
    rFlag.Right = ddsd2.lWidth


    'blt to the backbuffer from our  surface to
    'the screen surface such that our bitmap
    'appears over the window
    ddrval = backbuffer.BltFast(0, 0, flagsurf, rFlag, DDBLTFAST_WAIT)

    
    'Calculate the frame rate
    If i = 30 Then
        If tLast <> 0 Then fps = 30 / (Timer - tLast)
        tLast = Timer
        i = 0
    End If
    i = i + 1
    If Quit = False And Begin = True Then
    'set forecolor to lightgreen
    backbuffer.SetForeColor &HFFFFC0
    'draw the frame rate in the lower left corner
    Call backbuffer.DrawText(10, 460, "FPS: " + Format$(fps, "#.0"), False)
    'the all time high score
    Call backbuffer.DrawText(350, 112, "" + Format$(HighScore, "#0"), False)
    'set forecolor to red
    backbuffer.SetForeColor &HFF&
    'draw the ball number
    If Lost = False Then Call backbuffer.DrawText(554, 40, BallNumber, False)
    'if you lose then
    If Lost = True Then Call backbuffer.DrawText(284, 336, "Game Over", False)
    'loading another ball
    If LoadBall = True Then Call backbuffer.DrawText(286, 366, "Loading...", False)
    'if your score is higher than the highscore
    backbuffer.SetForeColor &HFFFFFF
    If Score > HighScore Then Call backbuffer.DrawText(287, 330, "High Score!", False)
    'players score
    backbuffer.SetForeColor &HFF&
    Call backbuffer.DrawText(300, 132, "" + Format$(Score, "#0"), False)
    ElseIf Quit = False And Begin = False Then
    backbuffer.SetForeColor vbWhite
    Call backbuffer.DrawText(252, 220, "Pinball in DirectDraw", False)
    Call backbuffer.DrawText(294, 240, "A demo.", False)
    BallNumber = 1
    GoTo finished:
    Else
    backbuffer.SetForeColor vbWhite
    If CreditA > 220 Then
    Call backbuffer.DrawText(260, CreditA, "Thanks for playing!", False)
    Call backbuffer.DrawText(110, CreditB, "Any comments on how to make this better, are always welcomed.", False)
    CreditA = CreditA - 1
    CreditB = CreditB - 1
    Else
    Call backbuffer.DrawText(260, CreditA, "Thanks for playing!", False)
    Call backbuffer.DrawText(110, CreditB, "Any comments on how to make this better, are always welcomed.", False)
    End If
    Quit = True
    GoTo finished:
    End If
    
    
    'calcultate wich frame# we are on in the sprite bitmap
    t2 = Timer
    If t <> 0 Then
        
        'keep tract of the picture frames
        RflipperFrame = RflipperFrame + (t2 - t) * 40
        If Rf = True Then
        If RflipperFrame > rows * cols - 1 Then RflipperFrame = 2 'was cols - 1
        Else
        If RflipperFrame > rows * cols - 2 Then RflipperFrame = 0 'was cols - 1
        End If
        LflipperFrame = LflipperFrame + (t2 - t) * 40
        If Lf = True Then
        If LflipperFrame > rows * cols - 1 Then LflipperFrame = 2 'was cols - 1
        Else
        If LflipperFrame > rows * cols - 2 Then LflipperFrame = 0 'was cols - 1
        End If
    End If
    t = t2
    DoEvents 'we need this so the timers run also
    'this controls the ball while
    'it is still in the shoot
    If Commit = False Then
    h1 = -2
    v1 = -1
    If Shoot = False Then
    rSprite5.Left = 430 'the balls resting position
    rSprite5.Top = 390
    x = 430
    y = 390
    g = 0.1
    d = 0 'added
    End If
    'we pushed the spacebar
    If Pull = True Then
    rSprite6.Left = 428 'the plunger
    rSprite6.Top = 400 + Plunger
    Else
    rSprite6.Left = 428
    rSprite6.Top = 400
    End If
    '
    rSprite2.Left = 429 'sprite cover over the plunger
    rSprite2.Top = 421
    '****************
    'this block of code controls the ball
    '****************
    If Shoot = True Then
    'you shot the ball
    If Point(x + 8, y + g - 2) = RGB(0, 128, 0) Then
    x = x - 1
    y = y - 2
    Commit = True
    Shoot = False
    Detain = False
    ElseIf Point(x + 11, y + g) = RGB(0, 0, 128) Then
    x = x - 1
    ElseIf Point(x + 1, y + g) = RGB(0, 0, 128) Then
    x = x + 2
    Else
    End If
    x = x - 0.09
    v = v + 0.1
    g = g + g * 0.07
    y = y - v
    rSprite5.Left = x
    rSprite5.Top = y + g
    If rSprite5.Top > 390 Then
    d = d - 0.4
    y = 390
    g = 1
    v = v + d
    If v < 0 Then
    v = 0
    d = 0
    g = 0.1
    Pull = False
    Shoot = False
    End If
    Else
    End If
    Else
    End If
    GoTo 11
    End If
    '
    '
    'the ball is in the playing field (past the corner)
    If Commit = True Then
    Timer2.Enabled = False
    rSprite2.Left = 429 'cover
    rSprite2.Top = 421
    If Point(x, y + 8) = RGB(0, 255, 255) And b = 0 Then
    b = 1
    h1 = ((Rnd * 6) + 4) / 10
    v1 = 0.3
    Else
    End If
    If Point(x + 8, y) = RGB(0, 128, 0) And b = 0 Then
    x = x - 0.5
    y = y + 1.5
    Else
    End If
    'if the ball gets shot back up by a bumper
    If Point(x + 8, y) = RGB(0, 128, 0) And Inplay = True And y < 230 Then
    x = x - 6
    y = y + 6
    h1 = 0
    v = 0
    v1 = 1.5
    d = 0
    Else
    End If
    'the troughs
    'bottom of ball
    If Point(x + 8, y + 15) = RGB(128, 128, 0) Then
    y = y - 4
    v1 = -0.4
    Else
    End If
    '
    If Point(x + 8, y + 15) = RGB(192, 192, 192) Then
    y = y - 4
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = -1.6
    Else
    End If
    If Point(x + 8, y) = RGB(128, 128, 0) Then
    y = y + 4
    v1 = 0.5
    Else
    End If
    'top of ball
    If Point(x + 8, y) = RGB(192, 192, 192) Then
    Randomize
    r = Int(Rnd * 2)
    If r = 0 Then
    x = x + 3
    y = y + 4
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = 1.6
    ElseIf r = 1 Then
    x = x - 3
    y = y + 4
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = 1.6
    Else
    End If
    End If
   
    'left side of ball
    If Point(x + 2, y + 8) = RGB(128, 128, 0) Then
    x = x + 1
    y = y + 2
    Else
    End If
    
    'left lower corner
    If Point(x + 4, y + 13) = RGB(192, 192, 192) Then
    x = x + 6
    y = y - 6
    'draw the lights when the ball hits any bumper
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = -1.6
    h1 = 2
    End If
    'right lower corner
    If Point(x + 12, y + 13) = RGB(192, 192, 192) Then
    x = x - 6
    y = y - 6
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = -1.6
    h1 = -2
    End If
    'left upper corner
    If Point(x + 4, y + 4) = RGB(192, 192, 192) Then
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = 1.6
    h1 = 2
    End If
    'right upper corner
    If Point(x + 12, y + 4) = RGB(192, 192, 192) Then
    Call backbuffer.DrawCircle(341, 275, 4)
    Call backbuffer.DrawCircle(379, 254, 4)
    Call backbuffer.DrawCircle(284, 257, 4)
    BumperHit.Play DSBPLAY_DEFAULT
    Score = Score + 100
    v1 = 1.6
    h1 = -2
    End If
    'right side of ball
    If Point(x + 14, y + 8) = RGB(128, 128, 0) Then
    x = x - 1
    y = y + 2
    Else
    End If
    If Point(x + 14, y + 8) = RGB(0, 0, 128) And Commit = True Then
    x = x - 6
    h1 = -0.5
    Else
    End If
    
    'ball hits trough trigger
    If Point(x + 8, y + 16) = RGB(0, 128, 128) Then
    Score = Score + 25
    Inplay = True
    'play the sound
    BonusHit.Play DSBPLAY_DEFAULT
    d = 0
    h1 = 0
    v1 = 1.6
    Else
    End If
    'ball hits sidewalls
    If y < 420 And x < 286 And Point(x + 1, y + 8) = RGB(128, 0, 0) Then
    x = x + 6
    h1 = 1.3
    d = 0
    ElseIf y < 355 And x > 286 And Point(x + 14, y + 8) = RGB(128, 0, 0) Then
    x = x - 6
    h1 = -1.3
    d = 0
    ElseIf Point(x + 2, y + 14) = RGB(128, 0, 0) Then
    x = x + (Rnd * 6) + 3
    y = y - 5
    h1 = 1.4
    v1 = -0.6
    d = 0
    Else
    End If
    'ball hits the grey bumper
    If Point(x + 14, y + 2) = RGB(128, 128, 128) Then
    v1 = 2
    h1 = -2
    ElseIf Point(x + 14, y + 14) = RGB(128, 128, 128) Then
    x = x - 3
    v1 = -0.2
    h1 = -1
    v = 0
    d = 0
    Else
    End If
    'this isn't optimized yet,
    'ball hits the flippers
    If Rf = True Then
    'your holding up the flipper to stop the ball
    If Point(x + 8, y + 16) = RGB(255, 128, 128) Then
    h1 = 0.1
    v1 = 0.1
    d = 0
    x = x + 1
    y = y - 1
    Detain = True
    Else
    End If
    End If
    If Lf = True Then
    If Point(x + 8, y + 16) = RGB(255, 128, 128) Then
    v1 = -3
    h1 = -1.5
    Else
    End If
    End If
    If Rf = True And Detain = True Then
    If Point(x + 8, y + 16) = RGB(255, 128, 128) Then
    v1 = -3
    h1 = -1.5
    Detain = False
    Else
    End If
    End If
    
    If Flipper = True And Detain = False Then
    If Point(x + 8, y + 15) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = -1.5
    ElseIf Point(x + 16, y + 16) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = -1.5
    ElseIf Point(x, y + 16) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    ElseIf Point(x + 12, y + 14) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    'if speed is high, catch the top
    ElseIf Point(x + 8, y) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    ElseIf Point(x + 4, y + 14) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    Else
    End If
    End If
    If Flipper = False And Detain = False Then
    If Point(x + 8, y + 15) = RGB(255, 0, 128) Then
    v1 = -1
    h1 = -1.5
    ElseIf Point(x + 16, y + 16) = RGB(255, 0, 128) Then
    v1 = -1
    h1 = -1.5
    ElseIf Point(x, y + 16) = RGB(255, 0, 128) Then
    v1 = -1
    h1 = 1.5
    ElseIf Point(x + 12, y + 14) = RGB(255, 0, 128) Then
    v1 = -1
    h1 = 1.5
    ElseIf Point(x + 4, y + 14) = RGB(255, 0, 128) Then
    v1 = -1
    h1 = 1.5
    ElseIf Point(x, y + 6) = RGB(255, 0, 128) Then
    v1 = -1
    h1 = 1.5
    Else
    End If
    End If
    If Flipper = True And Detain = True Then
    If Point(x + 8, y + 15) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = -1.5
    Detain = False
    ElseIf Point(x + 16, y + 16) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = -1.5
    Detain = False
    ElseIf Point(x, y + 16) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    Detain = False
    ElseIf Point(x + 12, y + 14) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    Detain = False
    ElseIf Point(x + 4, y + 14) = RGB(255, 0, 128) Then
    v1 = -3
    h1 = 1.5
    Detain = False
    Else
    End If
    End If

    'players score
    If Point(x + 8, y + 15) = RGB(0, 64, 64) Then
    Score = Score + 25
    Else
    End If
    'ball hits slider next to right flipper
    If Point(x + 8, y + 15) = RGB(0, 128, 0) And Detain = False Then
    y = y - 6
    v1 = -0.6
    h1 = 0.4
    ElseIf Point(x + 14, y + 14) = RGB(0, 128, 0) And Detain = True Then
    d = 0
    v1 = 0
    h1 = 0
    Else
    End If
    'ball hits yellow bumpers
    If Point(x + 8, y + 15) = RGB(255, 255, 128) Then
    Score = Score + 50
    v1 = -3
    h1 = (Rnd * 3) + 1.5
    'right side
    ElseIf Point(x + 15, y + 4) = RGB(255, 255, 128) Or Point(x + 14, y) = RGB(255, 255, 128) Then
    Score = Score + 50
    v1 = 1.5
    h1 = -2
    ElseIf Point(x + 16, y) = RGB(255, 255, 128) Or Point(x + 16, y + 8) = RGB(255, 255, 128) Then
    Score = Score + 50
    v1 = 1.5
    h1 = -2
    Else
    End If
    'ball hits purple bumper
    If Point(x + 14, y + 14) = RGB(255, 0, 255) Then
    Score = Score + 50
    v1 = -3
    h1 = -1 '-(Rnd * 3) + 1.5 (play with these!)
    Else
    End If
    'ball hits green ring, upper left corner
    If Point(x + 8, y) = RGB(0, 255, 0) Or Point(x + 16, y) = RGB(0, 255, 0) Then
    backbuffer.SetForeColor &HFFC0C0
    backbuffer.SetFillColor (&HFFC0C0)
    Call backbuffer.DrawCircle(242, 192, 6)
'play the sound
BonusHit.Play DSBPLAY_DEFAULT
    Score = Score + 500
    v1 = 2
    h1 = 2
    Else
    End If
    'ball hits the 2500pt green bumper
    If Point(x + 15, y + 8) = RGB(0, 255, 0) Then
    Score = Score + 2500
'play the sound
BonusHit.Play DSBPLAY_DEFAULT
    h1 = -1.5
    Else
    End If
    'maintain drag on the ball, one way of producing an arc motion
    If b = 1 Then d = d + 0.005
    x = x + h1 - d
    y = y + v1 + d
    Else
    End If
     'ball hits red ring (1000pts)
    If Point(x, y + 8) = RGB(255, 0, 0) And Ring = False Then
    v = 0
    h1 = 0
    v1 = 0
    d = 0
    g = 0
    b = 0
    x = 218
    y = 300
    Score = Score + 1000
'play the sound
BonusHit.Play DSBPLAY_DEFAULT
'ball delay for a few seconds to get your attention
    Timer3.Enabled = True
    GoTo 10
    'pop the ball back out on the table
    ElseIf Ring = True Then
    b = 1
    v = 1
    x = x + 6
    y = y + 6
    h1 = (Rnd * 2) + 1 'not always the same path
    v1 = 1
    Ring = False
    Else
    End If
    'you lose the ball; reset everything
    If y > 430 Then
    v = 0
    h1 = 0
    v1 = 0
    d = 0
    g = 0
    b = 0
    x = 315 'the ball hides below the flippers
    y = 430 'underneath the sprite tablecover
    Inplay = False
'play the sound
LoseBall.Play DSBPLAY_DEFAULT
'check the highscore
Open "highscore.dat" For Input As #1
Do While Not EOF(1)
Input #1, MyScore
Loop
Close #1
If Score > MyScore Then
Open "highscore.dat" For Output As #1
Write #1, Score
Close #1
Else
End If
If Score > 10000 And n = 0 Then
BallNumber = BallNumber - 1
n = n + 1
Else
End If
    BallNumber = BallNumber + 1
    If BallNumber = 4 Then
    Lost = True
    BallNumber = 1
    n = 0
    Else
    'pause for a few secs to allow you to cuss
    LoadBall = True
    Timer4.Enabled = True
    End If
    
    If Score > HighScore Then HighScore = Score
    Else
    End If
    
10  'finally, move the ball
    rSprite5.Left = x
    rSprite5.Top = y
11
    
    'blt to the backbuffer our sprites
    '********************************************************************************************************************************
    'this next line is important, we want
    'the ball to pass underneath all other
    'sprites so we draw it first, then the rest
    ddrval = backbuffer.BltFast(rSprite5.Left, rSprite5.Top, spritesurf1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'the ball
    '********************************************************************************************************************************
    ddrval = backbuffer.BltFast(rSprite6.Left, rSprite6.Top, spritesurf4, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'the shooter
    ddrval = backbuffer.BltFast(rSprite2.Left, rSprite2.Top, spritesurf3, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'the cover
    ddrval = backbuffer.BltFast(rSprite3.Left, rSprite3.Top, spritesurf2, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'anglecover
    ddrval = backbuffer.BltFast(rSprite7.Left, rSprite7.Top, spritesurf5, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'troughcover
    ddrval = backbuffer.BltFast(rSprite10.Left, rSprite10.Top, spritesurf8, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'tablehcover
    
    'only the flippers have picture frames
    'from the current frame select the bitmap we want to use
    'the right flipper
    col = RflipperFrame Mod cols
    row = Int(RflipperFrame / cols)
    rSprite.Left = col * spriteWidth
    rSprite.Top = row * spriteHeight
    rSprite.Right = rSprite.Left + spriteWidth
    rSprite.Bottom = rSprite.Top + spriteHeight
    'show the frame
    ddrval = backbuffer.BltFast(rSprite8.Left, rSprite8.Top, spritesurf6, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'rflipper
    'the left flipper
    col = LflipperFrame Mod cols
    row = Int(LflipperFrame / cols)
    rSprite.Left = col * spriteWidth
    rSprite.Top = row * spriteHeight
    rSprite.Right = rSprite.Left + spriteWidth
    rSprite.Bottom = rSprite.Top + spriteHeight
    'show the frame
    ddrval = backbuffer.BltFast(rSprite9.Left, rSprite9.Top, spritesurf7, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'lflipper
    
    'flip the back buffer to the screen
finished:
    primary.Flip Nothing, DDFLIP_WAIT

errOut:

End Sub

Sub EndIt()
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'until we begin
If Begin = False Then Exit Sub
'press the Esc key
If KeyCode = &H1B And Quit = False Then
Quit = True
finish
Exit Sub
'now any keystroke will end it
ElseIf Quit = True Then
'stop the music
pb = mciSendString("close " & "pinball.mid", 0&, 0, 0)
EndIt
Else
End If
'your pressing the spacebar
If KeyCode = &H20 And Shoot = False And LoadBall = False And Lost = False And Plunger < 12 Then
Plunger = Plunger + 1
If Plunger < 3 Then
v = 1
Else
v = Plunger / 3
End If
Pull = True
Else
End If
'you pressed the right flipper
If KeyCode = 76 Then Rf = True: Flipper = True
'...left flipper
If KeyCode = 65 Then Lf = True: Flipper = True
If KeyCode = &H77 And Lost = True Then
'new game reset
Score = 0
Commit = False
v = 0
h1 = 0
v1 = 0
d = 0
g = 0
b = 0
Lost = False
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'the spacebar
If KeyCode = &H20 Then Pull = False: Plunger = 0: Shoot = True
'L Key
If KeyCode = 76 Then Rf = False: Flipper = False
'A Key
If KeyCode = 65 Then Lf = False: Flipper = False
End Sub

Private Sub Form_Load()
'open and read the High Score
'file
On Error GoTo ErrorHandler:
Dim MyScore As Long
Open "highscore.dat" For Input As #1
Do While Not EOF(1)
Input #1, MyScore
Loop
Close #1
HighScore = MyScore
Init
Exit Sub
ErrorHandler:
MsgBox "Can't find the ""highscore.dat"" file, which should be in the same folder as the .exe.", vbCritical
Unload Me
End Sub

Private Sub Form_Paint()
blt
End Sub

Function ExModeActive() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = dd.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
    
End Function

Public Sub finish()
'here we flip from the opening screen to the game screen
Set flagsurf = Nothing
Set spritesurf = Nothing
    
   'load the bitmap into a surface
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = ddsd4.lWidth
    ddsd2.lHeight = ddsd4.lHeight
    If Begin = True And Quit = False Then
    Set flagsurf = dd.CreateSurfaceFromFile("pinball.bmp", ddsd2)
    Else
    'we are quitting the game:
    'to limit the file size, the credits screen
    'was not included but I'll leave the code here
    Set flagsurf = dd.CreateSurfaceFromFile("blank.bmp", ddsd2)
    End If
    
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
Begin = True
'switch the screen
pb = mciSendString("play " & "pinball.mid", 0&, 0, 0)
finish
'start the lights blinking
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
'draw all the blinking lights
Timer2.Interval = 200
backbuffer.SetForeColor &HFFC0FF
backbuffer.SetFillColor (&HFFC0FF)
Static lt As Integer
If lt = 1 Then
GoTo 1
ElseIf lt = 2 Then
GoTo 2
ElseIf lt = 3 Then
GoTo 3
ElseIf lt = 4 Then
GoTo 4
ElseIf lt = 5 Then
GoTo 5
Else
End If
Call backbuffer.DrawCircle(452, 393, 3)
Call backbuffer.DrawCircle(271, 204, 3)
Call backbuffer.DrawCircle(261, 338, 3)
lt = lt + 1
Exit Sub
1
Call backbuffer.DrawCircle(447, 351, 3)
Call backbuffer.DrawCircle(351, 204, 3)
Call backbuffer.DrawCircle(251, 330, 3)
lt = lt + 1
Exit Sub
2
Call backbuffer.DrawCircle(443, 311, 3)
Call backbuffer.DrawCircle(311, 204, 3)
Call backbuffer.DrawCircle(241, 322, 3)
lt = lt + 1
Exit Sub
3
Call backbuffer.DrawCircle(437, 268, 3)
Call backbuffer.DrawCircle(372, 204, 3)
lt = lt + 1
Exit Sub
4
Call backbuffer.DrawCircle(433, 225, 3)
Call backbuffer.DrawCircle(291, 204, 3)
lt = lt + 1
Exit Sub
5
Call backbuffer.DrawCircle(428, 182, 3)
Call backbuffer.DrawCircle(331, 204, 3)
lt = 0 'start over
End Sub


Private Sub Timer3_Timer()
'this is a small delay when the ball
'is in the 1000pt red ring
Timer3.Enabled = False
Ring = True
End Sub

Private Sub Timer4_Timer()
'a small delay after you lose a ball
Timer4.Enabled = False
Commit = False
LoadBall = False
Timer2.Enabled = True
End Sub


