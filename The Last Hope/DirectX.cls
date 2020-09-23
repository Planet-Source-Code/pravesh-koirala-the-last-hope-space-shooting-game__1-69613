VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Controls direct sound operations

Public Sub Initialize()
 Dim Buffer As DSBUFFERDESC
 Dim Format As WAVEFORMATEX

'On Local Error Resume Next
'Create Direct sound
Set DS = TheDx.DirectSoundCreate("")
If Err.Number <> 0 Then
MsgBox "Sorry You Don't Have DirectX installed", vbCritical
Exit Sub
End If
DS.SetCooperativeLevel frmMain.hWnd, DSSCL_NORMAL
'Create Direct sound Buffer
Buffer.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY
With Format
    .nFormatTag = WAVE_FORMAT_PCM
    .nChannels = 2
    .nBitsPerSample = 16
    .nBlockAlign = .nBitsPerSample \ 8 * .nChannels
    .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
End With
Set DB(0) = DS.CreateSoundBufferFromFile(App.Path + "\sound\Background.wav", Buffer, Format)
Set DB(1) = DS.CreateSoundBufferFromFile(App.Path + "\sound\Kaboom.wav", Buffer, Format)
Set DB(2) = DS.CreateSoundBufferFromFile(App.Path + "\sound\Kaboomex.wav", Buffer, Format)
Set DB(3) = DS.CreateSoundBufferFromFile(App.Path + "\sound\fire.wav", Buffer, Format)
Set DB(4) = DS.CreateSoundBufferFromFile(App.Path + "\sound\inventory.wav", Buffer, Format)
If MusicOff Then Exit Sub
DB(0).SetVolume -500
DB(0).Play DSBPLAY_LOOPING
End Sub

Public Sub Destroy()
Set TheDx = Nothing
Set DS = Nothing
End Sub

Public Sub PlayMusic(What As String)
If MusicOff Then Exit Sub
Select Case LCase(What)
Case "fire"
DB(3).Play DSBPLAY_DEFAULT
Case "explosionasteroid"
DB(2).Play DSBPLAY_DEFAULT
Case "explosionship"
DB(1).Play DSBPLAY_DEFAULT
Case "inventory"
DB(4).Play DSBPLAY_DEFAULT
End Select
End Sub
