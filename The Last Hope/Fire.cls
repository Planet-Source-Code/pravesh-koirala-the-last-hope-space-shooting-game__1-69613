VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Weapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class handles all operations regarding Fire like
'Creating fire
'Rotating it
'Drawing fire etc

Private FiresStatus(30) As Byte
Private FiresGraphicsDelay(30) As Integer
Dim Collision As cCollision
Public Sub DOFire()
'Check each fire for crossing the top limit
'Check each fire for colliding with asteroid
'Rotate each fire after a fixed time
'Finally paint the fire in MemDc
Set Collision = New cCollision
Dim i As Integer
    For i = 1 To 30
        If Fires(i) Then
            FiresY(i) = FiresY(i) - 5
            CheckForDistance (i)
            Collision.CheckCollisionAF i, FiresX(i), FiresY(i)
            FiresGraphicsDelay(i) = FiresGraphicsDelay(i) + 1
                If FiresGraphicsDelay(i) = 15 Then
                    FiresStatus(i) = Not (FiresStatus(i))
                    FiresGraphicsDelay(i) = 0
                End If
            PaintFire i, FiresStatus(i)
         End If
    Next
Set Collision = Nothing
End Sub

Public Sub CreateFire()
Dim DX As New DirectX
DX.PlayMusic "fire"
Set DX = Nothing
If FireQuality < 2 Then
        CurrFire = CurrFire + 1
        Fires(CurrFire) = 1
        FiresX(CurrFire) = ShipX + 31
        FiresY(CurrFire) = ShipY
Else
'Create 2 Fires
        CurrFire = CurrFire + 1
        Fires(CurrFire) = 1
        FiresX(CurrFire) = ShipX + 11
        FiresY(CurrFire) = ShipY

        CurrFire = CurrFire + 1
        Fires(CurrFire) = 1
        FiresX(CurrFire) = ShipX + 51
        FiresY(CurrFire) = ShipY
End If
End Sub

Private Sub PaintFire(Index As Integer, Status As Byte)
Dim HeightX As Integer, WidthX As Integer
HeightX = frmMain.picFire.Height
WidthX = frmMain.picFire.Width
Select Case Status
    Case 0
        BitBlt MemDc, FiresX(Index), FiresY(Index), WidthX \ 2, HeightX, _
        frmMain.maskFire.hdc, 0, 0, vbSrcAnd
        BitBlt MemDc, FiresX(Index), FiresY(Index), WidthX \ 2, HeightX, _
        frmMain.picFire.hdc, 0, 0, vbSrcPaint
    Case 255
        BitBlt MemDc, FiresX(Index), FiresY(Index), WidthX \ 2, HeightX, _
        frmMain.maskFire.hdc, 14, 0, vbSrcAnd
        BitBlt MemDc, FiresX(Index), FiresY(Index), WidthX \ 2, HeightX, _
         frmMain.picFire.hdc, 14, 0, vbSrcPaint
End Select
End Sub

Private Sub CheckForDistance(Index As Integer)
    If FiresY(Index) <= 0 Then
        ShiftFireNumber (Index)
    End If
End Sub

Private Sub ShiftFireNumber(Index As Integer)
Dim Collision As New cCollision
Collision.DestroyFire (Index)
Set Collision = Nothing
End Sub
