Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const SND_NOSTOP As Long = &H10

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public MemDc   As Long
Public MemBmp  As Long

'Variables for Fire
Public Fires(1 To 30)       As Byte
Public FiresX(1 To 30)      As Integer
Public FiresY(1 To 30)      As Integer
Public CurrFire             As Integer

'Variables for Asteroids
Public Asteroid(1 To 30)    As Byte
Public AsteroidX(1 To 30)   As Integer
Public AsteroidY(1 To 30)   As Integer
Public AsteroidRate(1 To 30) As Byte
Public CurrAsteroid         As Integer

'Variables for Explosion
Public Kaboom(1 To 30)      As Byte
Public KaboomX(1 To 30)     As Integer
Public KaboomY(1 To 30)     As Integer
Public KaboomS(1 To 30)     As Byte
Public CurrKaboom           As Byte

'Variables for ship's movement and ship's position
Public ShipX   As Integer
Public ShipY   As Integer
Public ShipSpeed As Integer
Public He      As Integer
Public Wi      As Integer

'Variables For Goodies
Public GoodyPresent As Boolean
Public GoodyX As Integer
Public GoodyY As Integer
Public GoodyT As Byte
Public GoodyTime As Byte

'Variables for GamePlay
Public Lives    As Byte
Public GameOver As Boolean
Public Score   As Long
Public Damage   As Byte
Public sState   As Integer
Public isShipDestroyed As Boolean
Public FireQuality As Byte
Public BombNumber As Byte
Public CheatOn As Boolean
Public GameTime As Integer

Public DX As New DirectX
Public TheDx As New DirectX7
Public DS As DirectSound
Public DB(4) As DirectSoundBuffer
Public MusicOff As Boolean

Public Sub Winner()
frmMain.Timer1.Enabled = False
frmMain.Timer2.Enabled = False
frmMain.Picture1.Left = 0
frmMain.Picture1.Top = 0
frmMain.Height = 317 * Screen.TwipsPerPixelY
frmMain.Width = 481 * Screen.TwipsPerPixelX
frmMain.Picture1.Visible = True
MsgBox "Earth is Proud Of you. You saved it!"
End Sub

Public Sub DoCheats()
frmMain.Timer1.Enabled = False
frmMain.Timer2.Enabled = False
Dim Cheatcode As String
Cheatcode = LCase(InputBox("Enter the CheatCode", "Cheat Console"))
Select Case Cheatcode
    Case "givemealife"
        Lives = Lives + 1
        Score = Score - 5000
    Case "lightspeed"
        ShipSpeed = ShipSpeed + 1
        Score = Score - 2000
    Case "hiroshima"
        BombNumber = BombNumber + 1
        Score = Score - 3000
    Case "qualityfire"
        FireQuality = FireQuality + 1
        Score = Score - 2000
End Select
frmMain.Timer1.Enabled = True
frmMain.Timer2.Enabled = True
End Sub
