VERSION 5.00
Begin VB.Form frmBball 
   Caption         =   "BouncingBall"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   Icon            =   "game1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdMin 
      Caption         =   "_"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "game1.frx":0ECA
      Top             =   0
      Width           =   6015
   End
   Begin VB.Timer timBall 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   1680
      Top             =   2640
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "start"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Shape shpBall 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "frmBball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BallX As Integer
Dim BallY As Integer
Dim BallDirx As Integer
Dim BallDiry As Integer

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
End Sub

Private Sub Form_Load()
 BallX = frmBball.ScaleWidth - shpBall.Width
 BallY = frmBball.ScaleHeight - shpBall.Height
 shpBall.Left = BallX
 shpBall.Top = BallY
 BallDiry = -1
 BallDirx = -1
 frmBball.BackColor = vbMagenta
 End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmStart.Show
End Sub

Private Sub timBall_Timer()
BallX = BallX + BallDirx * frmBball.ScaleWidth / 50
If BallX < 0 Then
  Beep
  BallX = 0
  BallDirx = 1
ElseIf BallX > frmBball.ScaleWidth - shpBall.Width Then
  Beep
  BallX = frmBball.ScaleWidth - shpBall.Width
  BallDirx = -1
End If
shpBall.Left = BallX
BallY = BallY + BallDiry * (frmBball.ScaleHeight - 615) / 50    '615 is height of text bar
If BallY < 615 Then
  Beep
  BallY = 615
  BallDiry = 1
ElseIf BallY > (frmBball.ScaleHeight) - shpBall.Height Then
  Beep
  BallY = (frmBball.ScaleHeight) - shpBall.Height
  BallDiry = -1
End If
shpBall.Top = BallY
End Sub
