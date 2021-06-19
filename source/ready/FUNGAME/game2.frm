VERSION 5.00
Begin VB.Form frmField 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit The Ball"
   ClientHeight    =   8475
   ClientLeft      =   555
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "game2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9450
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   8160
      Width           =   495
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "_"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   $"game2.frx":0ECA
      Top             =   8160
      Width           =   8295
   End
   Begin VB.Timer timBall 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   2640
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "start"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image imgRod 
      Height          =   150
      Left            =   5160
      Picture         =   "game2.frx":0F3B
      Top             =   7800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Shape shpBall 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BallX As Integer
Dim BallY As Integer
Dim BallDirx As Integer
Dim BallDiry As Integer

Private Sub cmdCredit_Click()
frmCredit.Visible = True
frmField.Visible = False
End Sub

Private Sub cmdMin_Click()
If timBall.Interval < 400 Then
timBall.Interval = timBall.Interval + 5
Else
timBall.Interval = 70
End If
End Sub

Private Sub cmdPlus_Click()
If timBall.Interval > 5 Then
timBall.Interval = timBall.Interval - 5
Else
timBall.Interval = 10
End If
End Sub

Private Sub cmdStart_Click()
timBall.Enabled = True
cmdStart.Visible = False
shpBall.Visible = True
imgRod.Visible = True
frmField.Width = 10000
frmField.Height = 9000
End Sub

Private Sub Form_Load()
 BallX = frmField.ScaleWidth - shpBall.Width
 BallY = frmField.ScaleHeight - shpBall.Height
 shpBall.Left = BallX
 shpBall.Top = BallY
 BallDiry = -1
 BallDirx = -1
 frmField.BackColor = vbMagenta
 End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmStart.Show
End Sub

Private Sub timBall_Timer()
BallX = BallX + BallDirx * frmField.ScaleWidth / 50
If BallX < 0 Then
  Beep
  BallX = 0
  BallDirx = 1
ElseIf BallX > frmField.ScaleWidth - shpBall.Width Then
  Beep
  BallX = frmField.ScaleWidth - shpBall.Width
  BallDirx = -1
End If
shpBall.Left = BallX
BallY = BallY + BallDiry * frmField.ScaleHeight / 50
If BallY < 0 Then
  Beep
  BallY = 0
  BallDiry = 1
'ElseIf ((BallY \ 7425) * ((imgRod.Left + 1) \ (shpBall.Left + 1)) * ((imgRod.Left + 1051) \ (shpBall.Left + 1))) = 1 Then
 ' Beep
  'BallY = imgRod.Top - shpBall.Height
  'BallDiry = -1
ElseIf shpBall.Top > 7430 Then
    If shpBall.Top < 8000 Then
       If shpBall.Left > imgRod.Left Then
          If shpBall.Left < (imgRod.Left + 1050) Then
           Beep
           BallY = imgRod.Top - shpBall.Height
           BallDiry = -1
          End If
       End If
      Else
       BallY = 0
       BallDiry = 1
       Beep
     End If
ElseIf BallY > frmField.ScaleHeight - shpBall.Height Then
Beep
BallY = frmField.ScaleHeight - shpBall.Height
BallDiry = -1

  'ElseIf shpBall.Top + ShpBall.Height = imgRod.Top
End If
shpBall.Top = BallY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 11090 Then
imgRod.Left = X - 200

End If
End Sub

Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
'txtDown.Text = Str(KeyCode)
'If KeyCode = 37 Then
'Beep
''imgRod.Left = imgRod.Left - 400
'ElseIf KeyCode = 39 Then
'Beep
'imgRod.Left = imgRod.Left + 400
'End If

End Sub

