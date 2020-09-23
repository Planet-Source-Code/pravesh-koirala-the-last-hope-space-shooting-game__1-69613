VERSION 5.00
Begin VB.Form frmSetting 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Last Hope - Settings"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Text            =   "300"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "View High Score"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play Game>>>"
      Height          =   855
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disable Music"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Controls"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BackGround Story"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   480
      Picture         =   "frmSetting.frx":030A
      ScaleHeight     =   855
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   240
      Width           =   4620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Set Game Length (seconds)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
MsgBox "Earth is in Danger. Asteroids are Falling" & vbCrLf _
& "You are odered to hold Asteroids until Scientists" & vbCrLf _
& "find any Solution. You are THE LAST HOPE.", vbInformation + vbOKOnly
End Sub

Private Sub Command2_Click()
MsgBox "  Arrow Keys        --- Move" & vbCrLf _
& "  Control button   --- Fire" & vbCrLf _
& "  Space Bar          --- Fire bomb " & vbCrLf _
& "  Esc                     --- Pause game", vbInformation + vbOKOnly
End Sub

Private Sub Command3_Click()
Static Status As Boolean
If Not Status Then
MusicOff = True
MsgBox "Music Off"
Command3.Caption = "Enable Music"
Else
MusicOff = False
MsgBox "Music On"
Command3.Caption = "Disable Music"
End If
Status = Not Status
End Sub

Private Sub Command4_Click()
If Val(Text1.Text) < 100 Or Val(Text1.Text) > 900 Then
MsgBox "Time must be > 100 and < 900", vbCritical + vbOKOnly
Exit Sub
Else
GameTime = Val(Text1.Text)
End If
Unload Me
frmMain.Show
End Sub

Private Sub Command5_Click()
names = GetSetting("The Last Hope", "HS", "Name", "Pravesh")
sc = GetSetting("The Last Hope", "HS", "Score", "13000")
MsgBox "High score is " & sc & " By" & Chr(32) & names, vbInformation + vbOKOnly, "High Score"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static Cheat As String
Cheat = Cheat + Chr(KeyCode)
If Cheat = "ASTALAVISTA" Then
CheatOn = True
End If
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
Dim control As Variant
For Each control In Me
    If TypeOf control Is CommandButton Then
        control.BackColor = Me.BackColor
        control.FontSize = 10
        control.FontBold = True
    End If
Next control
End Sub
