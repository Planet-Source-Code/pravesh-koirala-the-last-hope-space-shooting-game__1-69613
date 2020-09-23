VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsteroids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class handles all actions regarding the movement of Asteroid
'Creating new Asteroids
'Their collision with Fire and Ship etc
'Drawing Asteroid in the MEMDC
'For asteroid's rotation
Private AsteroidS(0 To 30) As Byte
'WE shall be using i in many loops
Dim i As Byte

Public Sub DoAsteroids()
Static TimeDelay As Byte
'Delay the time so as not to refresh the asteroid state quickly
    TimeDelay = TimeDelay + 1
    If TimeDelay > 8 Then
        TimeDelay = 0
        ChangeAsteroidState
    End If
'Check collision and Draw all the asteroids
CheckCurrAsteroids
End Sub

Public Sub InitAsteroids()
'(Optional) Can be used to initialize game with few asteroids
    Randomize
    For i = 1 To 6
        Asteroid(i) = 1
        AsteroidX(i) = Int(Rnd * He + 1)
               AsteroidY(i) = Int(Rnd * 200 + 1)
        AsteroidS(i) = Int(Rnd * 10 + 1)
        AsteroidRate(i) = 1
        CurrAsteroid = CurrAsteroid + 1
    Next i
End Sub

Public Sub CreateAsteroid()
'Creates a new asteroid in every fixed interval of time triggered by timer2 in
'frmMain
Randomize
Dim i As Integer, Rate As Byte
i = CurrAsteroid + 1
        Asteroid(i) = 1
        Randomize
        AsteroidX(i) = Int(Rnd * He + 1)
        AsteroidY(i) = 0
        AsteroidS(i) = Int(Rnd * 10 + 1)
        Rate = Int(Rnd * 15 + 1)
        If Rate < 2 Then
        AsteroidRate(i) = 2
        Else
        AsteroidRate(i) = 1: End If
        CurrAsteroid = CurrAsteroid + 1
    
End Sub

Private Sub CheckCurrAsteroids()
'Check for all current asteroid and detect collision and others
    For i = 1 To CurrAsteroid
        If Asteroid(i) = 1 Then
            AsteroidY(i) = AsteroidY(i) + AsteroidRate(i)
            If AsteroidY(i) > (He - 50) Then
                'Asteroid that move fastly do more damage
                If AsteroidRate(i) = 1 Then
                Damage = Damage + 1
                Else
                Damage = Damage + 2
                End If
                'If more than 20 Asteroids pass through then game over
                If Damage = 20 Then GameOver = True
                DeleteAsteroid (i)
            End If
            DrawAsteroid i, AsteroidS(i - 1)
        End If
    Next
End Sub

Private Sub ChangeAsteroidState()
'Handles the Asteroid's rotation
    If CurrAsteroid = 0 Then Exit Sub
    For i = 0 To CurrAsteroid - 1
        If AsteroidS(i) > 8 Then AsteroidS(i) = 0
        AsteroidS(i) = AsteroidS(i) + 1
    Next
End Sub
Private Sub DrawAsteroid(ByVal Index As Integer, State As Byte)
    Dim Height As Integer, Xcord As Integer
    Height = frmMain.picAsteroids.Height
    Xcord = 50 * State
    BitBlt MemDc, AsteroidX(Index), AsteroidY(Index), 50, Height, frmMain.maskAsteroid.hdc, Xcord, 0, vbSrcAnd
    BitBlt MemDc, AsteroidX(Index), AsteroidY(Index), 50, Height, frmMain.picAsteroids.hdc, Xcord, 0, vbSrcPaint
End Sub


Private Sub DeleteAsteroid(Index As Integer)
'Destroys asteroid after Collision
    Dim Collision As New cCollision
    Collision.DestroyAsteroid (Index)
    Set Collision = Nothing
End Sub
