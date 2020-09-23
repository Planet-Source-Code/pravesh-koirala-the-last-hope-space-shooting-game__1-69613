VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCollision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This Class Handles all operations regarding Object collision like
'Checking collision of Asteroid with Fire
'Checking collision of Asteroid with Ship
'Exploding Asteroid and Ship
'Deleting their values from memory

Dim i As Integer
'************************************Collision Module*****************************
Public Sub CheckCollisionAF(ByVal FireIndex As Integer, ByVal Xcord As Integer, ByVal Ycord As Integer)
'Check Collision of Fire with Asteroid
    
    For i = 1 To CurrAsteroid
        If Xcord > (AsteroidX(i) - 10) And Xcord < (AsteroidX(i) + 50) Then
             If Ycord < (AsteroidY(i) + 50) Then
               'Fast moving asteroid gives more point and goodies
                
                If AsteroidRate(i) = 2 Then
                    Score = Score + 200
                    Dim Goody As New cGoody
                    Goody.CreateGoody AsteroidX(i), AsteroidY(i)
                    Set Goody = Nothing
                Else
                    Score = Score + 100
                End If
                
               Call ExplodeAsteroid(FireIndex, i)
             End If
        End If
    Next
End Sub

Public Sub CheckCollisionAS(Xcord As Integer, Ycord As Integer)
'Checks collision of Asteroid with ship
'If ship is currently destroyed then Bye!!
    If isShipDestroyed Then Exit Sub
    For i = 1 To CurrAsteroid
            If (AsteroidX(i) + 50) > ShipX And (AsteroidX(i)) < (ShipX + 75) Then
                If (AsteroidY(i) + 50) > ShipY And (AsteroidY(i)) < (ShipY + 62) Then
                    ExplodeShip (i)
                End If
            End If
    Next i
End Sub

Public Function CheckCollisionGS(ByVal Xcord As Integer, ByVal Ycord As Integer) As Boolean
'Checks collision of Goodies with Ship
If Xcord > ShipX And Xcord < (ShipX + 79) Then
    If Ycord > ShipY And Ycord < (ShipY + 62) Then
        CheckCollisionGS = True
        DX.PlayMusic "Inventory"
    Else
        CheckCollisionGS = False
    End If
    
End If
End Function
Private Sub ExplodeAsteroid(FireIndex As Integer, AsteroidIndex As Integer)
'Destroy fire and Asteroid
'If weapon is upgraded then don't destroy fire
    If FireQuality Mod 2 = 0 Then DestroyFire (FireIndex)
    DestroyAsteroid (AsteroidIndex)
End Sub

Private Sub ExplodeShip(AsteroidIndex As Integer)
'Destroy both ship and Asteroid
    DestroyAsteroid (AsteroidIndex)
    NewExplosion ShipX + 14, ShipY + 6
    DX.PlayMusic "ExplosionShip"
    isShipDestroyed = True
    Lives = Lives - 1
    FireQuality = 0
    ShipSpeed = 2
    'So that it could not collide
    ShipX = -200
    ShipY = -200
    If Lives = 0 Then GameOver = True
End Sub

Public Sub DestroyFire(Index As Integer)
'Quite simple algorithm to delete Fire value
    If CurrFire = 0 Then Exit Sub
    If Index = CurrFire Then
        Fires(Index) = 0
        CurrFire = CurrFire - 1
        Exit Sub
    Exit Sub
    End If
    FiresX(Index) = FiresX(CurrFire)
    FiresY(Index) = FiresY(CurrFire)
    Fires(CurrFire) = 0
    CurrFire = CurrFire - 1
End Sub

Public Sub DestroyAsteroid(Index As Integer)
    NewExplosion AsteroidX(Index), AsteroidY(Index)
    
    If CurrAsteroid = 0 Then Exit Sub
    If Index = CurrAsteroid Then
    Asteroid(Index) = 0
    CurrAsteroid = CurrAsteroid - 1
    Exit Sub
    End If
    AsteroidX(Index) = AsteroidX(CurrAsteroid)
    AsteroidY(Index) = AsteroidY(CurrAsteroid)
    Asteroid(CurrAsteroid) = 0
    AsteroidRate(Index) = AsteroidRate(CurrAsteroid)
    CurrAsteroid = CurrAsteroid - 1
    DX.PlayMusic "ExplosionAsteroid"
End Sub

'********************************EXplosion Module**********************************

Public Sub DoExplosion()
'Called from main timer in frmMain
'This frequently updates the Explosion state
    For i = 1 To CurrKaboom
        KaboomS(i) = KaboomS(i) + 1
        If KaboomS(i) > 6 Then
            DeleteKaboom
        End If
        DrawKaboom i, KaboomS(i)
    Next
End Sub

Private Sub NewExplosion(Xcord As Integer, Ycord As Integer)
    CurrKaboom = CurrKaboom + 1
    Kaboom(CurrKaboom) = 1
    KaboomX(CurrKaboom) = Xcord
    KaboomY(CurrKaboom) = Ycord
    KaboomS(CurrKaboom) = 0
End Sub

Private Sub DrawKaboom(ByVal Index As Integer, ByVal State As Integer)
'Draw Explosion in Memdc
Dim Xcord As Integer
Xcord = 50 * State
BitBlt MemDc, KaboomX(Index), KaboomY(Index), 50, 50, frmMain.maskBang.hdc, Xcord, 0, vbSrcAnd
BitBlt MemDc, KaboomX(Index), KaboomY(Index), 50, 50, frmMain.picBang.hdc, Xcord, 0, vbSrcPaint
End Sub

Private Sub DeleteKaboom()
'Delete the Last Explosion
'Based in the algorithm, the explosion that starts first,ends first
For i = 1 To CurrKaboom - 1
    KaboomX(i) = KaboomX(i + 1)
    KaboomY(i) = KaboomY(i + 1)
Next
CurrKaboom = CurrKaboom - 1
End Sub

