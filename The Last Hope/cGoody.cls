VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGoody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This Class Controls all operations regarding Goodies like
'Creating and determining the type of goodies
'Implementing their Effects etc


Public Sub DoGoody()
Dim Collision As New cCollision
If GoodyTime > 5 Then
GoodyPresent = False
End If
    DisplayGoody
    If Collision.CheckCollisionGS(GoodyX, GoodyY) Then
    Score = Score + 500
    UseGoody
    GoodyPresent = False
End If
Set Collision = Nothing
End Sub

Public Function CreateGoody(ByVal Xcord As Integer, ByVal Ycord As Integer) As String
Randomize
GoodyPresent = True
GoodyX = Xcord
GoodyY = Ycord
GoodyT = Int(Rnd * 4 + 1)
CreateGoody = TranslateGoody()
GoodyTime = 0
End Function

Private Function TranslateGoody() As String
Select Case GoodyT
    Case 1
    TranslateGoody = "S"
    Case 2
    TranslateGoody = "W"
    Case 3
    TranslateGoody = "L"
    Case 4
    TranslateGoody = "B"
End Select
End Function

Private Sub UseGoody()
Select Case TranslateGoody
Case "S"
If ShipSpeed < 5 Then ShipSpeed = ShipSpeed + 1
Case "W"
If FireQuality < 3 Then FireQuality = FireQuality + 1
Case "L"
If Lives < 6 Then Lives = Lives + 1
Case "B"
BombNumber = BombNumber + 1
End Select
End Sub

Private Sub DisplayGoody()
SetTextColor MemDc, vbGreen
TextOut MemDc, GoodyX, GoodyY, TranslateGoody(), 1
SetTextColor MemDc, vbWhite
End Sub
