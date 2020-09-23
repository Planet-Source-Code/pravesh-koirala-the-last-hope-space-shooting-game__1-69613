VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "300"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   744
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Left            =   8040
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   11
      Top             =   9960
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.PictureBox picDamage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   9480
      Picture         =   "frmMain.frx":71184
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   10
      Top             =   8400
      Width           =   1200
   End
   Begin VB.PictureBox picLife 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   9000
      Picture         =   "frmMain.frx":71946
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   9
      Top             =   8280
      Width           =   300
   End
   Begin VB.PictureBox picBkGround 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   12720
      Picture         =   "frmMain.frx":71D0C
      ScaleHeight     =   1395
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   12960
      Top             =   6360
   End
   Begin VB.PictureBox maskBang 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   7080
      Picture         =   "frmMain.frx":15A62A
      ScaleHeight     =   750
      ScaleWidth      =   5250
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox picBang 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   6720
      Picture         =   "frmMain.frx":1673E6
      ScaleHeight     =   750
      ScaleWidth      =   5250
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox maskAsteroid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   6720
      Picture         =   "frmMain.frx":1741A2
      ScaleHeight     =   750
      ScaleWidth      =   7500
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox picAsteroids 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   6720
      Picture         =   "frmMain.frx":1866DE
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox maskFire 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13320
      Picture         =   "frmMain.frx":198C1A
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   12840
      Picture         =   "frmMain.frx":19931C
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   12480
      Top             =   6360
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   6720
      Picture         =   "frmMain.frx":1998F2
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   559
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8385
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   6720
      Picture         =   "frmMain.frx":1CC6F4
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   559
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   8385
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FOR EXPERIENCED GAME PROGRAMMERS IN VB6:-
'Please Read the Advanced section under ReadMe.txt file included with this Game
'This is my first game in VB or infact in any PG Language
'Its not the final release of this game as I plan to extend it
'Game is updated with the help of two timers 1 and 2

'Declarations to the Classes
Dim Fire As Weapon
Dim Asteroid As CAsteroids
Dim Explosion As cCollision
Dim Goody As cGoody

'Gameplay variables
Dim Keypressed As Boolean
Dim TimeVal As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static Cheatcode As String
If isShipDestroyed Then Exit Sub
'Control fires and space releases Bomb
    If KeyCode = vbKeyControl Then
        If Not Keypressed Then
            If Score > 10 Then Score = Score - 20
            Fire.CreateFire
            Keypressed = True
        End If
    ElseIf KeyCode = vbKeySpace Then
        If BombNumber Then Destroyall
    ElseIf KeyCode = 192 And CheatOn Then DoCheats
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Timer2.Enabled = Not Timer2.Enabled
Timer1.Enabled = Not Timer1.Enabled
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Just only allow the player to fire once with Control
If KeyCode = vbKeyControl Then
Keypressed = False
End If
End Sub

Private Sub Form_Load()

'Set the height and width of the Form
Me.Height = Screen.TwipsPerPixelX * 500
Me.Width = Screen.TwipsPerPixelY * 600
He = Me.ScaleHeight
Wi = Me.ScaleWidth

'PlayMusic "Background"
'Set the Gameplay variables
Lives = 3
Damage = 0
CurrFire = 0
ShipSpeed = 2
BombNumber = 1
FireQuality = 0

'Set the ship position at the centre and States
ShipX = Me.ScaleWidth \ 2 - 38
ShipY = Me.ScaleHeight - picShip.Height
sState = 0
DrawShip Me.hdc, ShipX, ShipY, 3

'Pass all the keydown events to Form
Me.KeyPreview = True

'Create New Instance of Classes
Set Fire = New Weapon
Set Asteroid = New CAsteroids
Set Explosion = New cCollision
Set Goody = New cGoody
DX.Initialize
'Asteroid.InitAsteroids
'Create a compatible bitmap and DC and avoid Flicker
MemDc = CreateCompatibleDC(Me.hdc)
MemBmp = CreateCompatibleBitmap(Me.hdc, Me.ScaleWidth, Me.ScaleHeight)
DeleteObject SelectObject(MemDc, MemBmp)
SetBkMode MemDc, 0
SetTextColor MemDc, vbWhite
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim names As String
'Delete the DC and Bitmap
    DeleteObject MemBmp
    DeleteDC MemDc
'Destroy the instance of classes
    Set Fire = Nothing
    Set Asteroid = Nothing
    Set Explosion = Nothing
    Set Goody = Nothing
    Set DX = Nothing
    
    If Score > Val(GetSetting("The Last Hope", "HS", "Score")) Then
        names = InputBox("You made a High Score" & vbCrLf & "Enter Your Name" _
        , "Congratulations")
        If names <> "" Then
            SaveSetting "The Last Hope", "HS", "Name", names
            SaveSetting "The Last Hope", "HS", "Score", Score
        End If
    End If
End
End Sub

Private Sub DrawShip(Dc As Long, x As Integer, y As Integer, Index As Integer)
'We shall Draw the Ship only if it is not Destroyed
'Else we shall delay the time and display a new ship at Centre
Dim Xcord As Integer, Ycord As Integer
Static ShipDisplayDelay As Integer

If isShipDestroyed Then
    ShipDisplayDelay = ShipDisplayDelay + 1
        If ShipDisplayDelay > 60 Then
            ShipX = Me.ScaleWidth \ 2 - 38
            ShipY = Me.ScaleHeight - picShip.Height
            sState = 0
            ShipDisplayDelay = 0
            isShipDestroyed = False
        End If
    Exit Sub
End If
'The ship position in the picture is quite Odd so we shall use tricks to draw ship
'correctly
If Index < 0 Then
Xcord = (Index + 6) * 79
Ycord = picShip.ScaleHeight \ 2
Else
Xcord = Index * 79
Ycord = 0
End If
BitBlt Dc, x, y, 79, picMask.ScaleHeight \ 2, picMask.hdc, Xcord, Ycord, vbSrcAnd
BitBlt Dc, x, y, 79, picShip.ScaleHeight \ 2, picShip.hdc, Xcord, Ycord, vbSrcPaint
End Sub

Private Sub CheckForKeys()
'Actually I wanted the ship to fire as well as move simultaneously so I used both
'API and keydown seperately.

Static ShipState1 As Integer, ShipState2 As Integer
'The noKEys value is used to auto straighten the ship when movement keys are not
'pressed.
Dim noKEys As Boolean
Static TimeDelay As Integer
noKEys = True
If isShipDestroyed Then Exit Sub

If GetAsyncKeyState(vbKeyUp) Then
    If ShipY > 0 Then ShipY = ShipY - ShipSpeed
    noKEys = False
End If

If GetAsyncKeyState(vbKeyDown) Then
    If ShipY < (Me.ScaleHeight - picShip.ScaleHeight \ 2) Then ShipY = ShipY + ShipSpeed
    noKEys = False
End If

If GetAsyncKeyState(vbKeyRight) Then
    If ShipX < (Me.ScaleWidth - 79) Then ShipX = ShipX + ShipSpeed
    ShipState1 = ShipState1 + 2
    If ShipState1 > 2 Then
    If sState < 6 Then sState = sState + 1
    ShipState1 = 0
    End If
noKEys = False
End If

If GetAsyncKeyState(vbKeyLeft) Then
    If ShipX > 0 Then ShipX = ShipX - ShipSpeed
    ShipState2 = ShipState2 + 1
    If ShipState2 > 2 Then
    If sState > -5 Then sState = sState - 1
    ShipState2 = 0
    End If
    noKEys = False
End If

'Check for Collisiom of Ship with Asteroid
Dim Collision As New cCollision
    Collision.CheckCollisionAS ShipX, ShipY
If noKEys Then
  TimeDelay = TimeDelay + 1
    If TimeDelay > 5 And sState <> 0 Then
        TimeDelay = 0
        If sState < 0 Then sState = sState + 1
        If sState > 0 Then sState = sState - 1
    End If
TimeDelay = TimeDelay + 1
End If
Set Collision = Nothing
End Sub

Private Sub PaintItBlack(Dc As Long)
'Repaint the BackGround each Time.
Dim hBrush As Long, blah As RECT
hBrush = CreatePatternBrush(picBkGround.Picture)
With blah
    .Top = 0
    .Left = 0
    .Bottom = He
    .Right = Wi
End With
FillRect Dc, blah, hBrush
DeleteObject hBrush

'Draw the lives in the Top Corner
Dim i As Integer, Startpos As Integer
Startpos = Me.ScaleWidth - (Lives - 1) * 20
For i = 1 To Lives - 1
    BitBlt MemDc, Startpos, 0, 20, 15, picLife.hdc, 0, 0, vbSrcCopy
    Startpos = Startpos + 20
Next
If CheatOn Then SetTextColor MemDc, vbRed
'Draw Damage in Top Left
    BitBlt MemDc, 0, 0, 4 * (20 - Damage), 8, picDamage.hdc, 0, 0, vbSrcCopy
    TextOut MemDc, 10, 10, "Damage " & Damage & Chr(32), 9
    TextOut MemDc, Me.ScaleWidth \ 2 - 20, 0, "Score " & Format(Score, "0000000"), 13
'Draw Bombnumbers
    TextOut MemDc, Wi \ 2 - 10, 15, BombNumber & " Bombs Left", 12
End Sub

Private Sub CopyStuffs()
'Copy all the stuffs of the Memory Dc into Our Form
BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, MemDc, 0, 0, vbSrcCopy
End Sub

Private Sub CheckFires()
'CHeck and Update the State of Fires
If CurrFire Then Fire.DOFire
End Sub

Private Sub CheckAsteroids()
'Check and Update the State of Asteroids
Asteroid.DoAsteroids
End Sub

Private Sub CheckExplosion()
'Check and Update the State of Explosions
If CurrKaboom Then Explosion.DoExplosion
End Sub

Private Sub CheckGoody()
'Check if Extra Goodies are available or not
If GoodyPresent Then Goody.DoGoody
End Sub

Private Function TranslateColor(aColor As OLE_COLOR) As Long
'I don't quite know about this. I got it from elsewhere.
    Dim newcolor As Long
    OleTranslateColor aColor, Me.Palette, newcolor
    TranslateColor = newcolor
End Function

Private Sub Destroyall()
'If you Have Bomb left and you release it then Destroy all Current Asteroids
'But it won't provide you points
Dim i As Integer
For i = 1 To CurrAsteroid
Explosion.DestroyAsteroid i
Next
BombNumber = BombNumber - 1
End Sub

Private Sub Timer1_Timer()
'This is a secondary timer.
'It creates a new Asteroid at a regular interval and also decides the Goody Time
            Timer1.Enabled = False
            Timer2.Enabled = False
            Asteroid.CreateAsteroid
            Timer1.Enabled = True
            Timer2.Enabled = True
            TimeVal = TimeVal + 1
            Me.Caption = GameTime - TimeVal & " seconds remaining"
            If TimeVal = GameTime Then Winner
            GoodyTime = GoodyTime + 1
End Sub

Private Sub Timer2_Timer()
'This is the main timer that updates all contents in the screen
'I heard somewhere that a Loop is good than Timer but I was quite not sure
'If anybody wants to help then please help
If GameOver Then
MsgBox "Game over"
End
End If
PaintItBlack MemDc
            DrawShip MemDc, ShipX, ShipY, sState
            CheckForKeys
            CheckFires
            CheckAsteroids
            CheckExplosion
            CheckGoody
            CopyStuffs
End Sub

