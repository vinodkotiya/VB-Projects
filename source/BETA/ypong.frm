VERSION 5.00
Begin VB.Form frmField 
   BackColor       =   &H00C0FFFF&
   Caption         =   "VINTRACK (DEMO VER.)"
   ClientHeight    =   8190
   ClientLeft      =   2100
   ClientTop       =   945
   ClientWidth     =   11880
   Icon            =   "ypong.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtP2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   10080
      TabIndex        =   4
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtP1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   615
      Width           =   615
   End
   Begin VB.PictureBox picRit 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11835
      TabIndex        =   2
      Top             =   7560
      Width           =   11895
   End
   Begin VB.CommandButton cmdCredit 
      BackColor       =   &H00FFFF80&
      Caption         =   "Credit"
      Height          =   1095
      Left            =   2880
      Picture         =   "ypong.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtDown 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Text            =   "START THE GAME"
      Top             =   3360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer timBall 
      Enabled         =   0   'False
      Interval        =   170
      Left            =   600
      Top             =   3360
   End
   Begin VB.Image imgTwo2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "ypong.frx":1D94
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgTwo1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "ypong.frx":2C66
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgThree 
      Height          =   330
      Left            =   6960
      Picture         =   "ypong.frx":3B38
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgResult 
      Height          =   750
      Left            =   3840
      Picture         =   "ypong.frx":41E2
      Top             =   3000
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgVin 
      Height          =   750
      Left            =   480
      Picture         =   "ypong.frx":84A4
      Top             =   2160
      Visible         =   0   'False
      Width           =   10800
   End
   Begin VB.Image imgTwo 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   6960
      Picture         =   "ypong.frx":11586
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgOne2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "ypong.frx":11C30
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgOne1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "ypong.frx":12B02
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgNull2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "ypong.frx":139D4
      Top             =   120
      Width           =   840
   End
   Begin VB.Image imgNull1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "ypong.frx":148A6
      Top             =   120
      Width           =   840
   End
   Begin VB.Image imgBot 
      Height          =   1200
      Left            =   0
      Picture         =   "ypong.frx":15778
      Top             =   6360
      Width           =   12000
   End
   Begin VB.Image imgOne 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   6960
      Picture         =   "ypong.frx":255BA
      Top             =   360
      Width           =   435
   End
   Begin VB.Image imgCap 
      Height          =   1050
      Left            =   0
      Picture         =   "ypong.frx":25C64
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image imgP1 
      Height          =   1920
      Left            =   1320
      Picture         =   "ypong.frx":33B66
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgP2 
      Height          =   1920
      Left            =   10080
      Picture         =   "ypong.frx":347A8
      Top             =   4680
      Width           =   240
   End
   Begin VB.Shape shpBall 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C0C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   375
      Left            =   360
      Shape           =   3  'Circle
      Top             =   2760
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
Dim BallX As Integer, Bot As Integer
Dim BallY As Integer, Tim As Integer
Dim BallDirx As Integer, Strok As Integer
Dim BallDiry As Integer, Cap As Integer, Div As Integer
Dim Sp1 As Integer, Sp2 As Integer
Dim Vin1 As Integer, Vin2 As Integer

Private Sub cmdCredit_Click()
frmCredit.Visible = True
frmField.Visible = True
'timBall.Interval = 0
End Sub



Private Sub Form_Load()
 BallX = frmField.ScaleWidth / 2
 BallY = frmField.ScaleHeight / 2
 shpBall.Left = BallX
 shpBall.Top = BallY
 BallDiry = 1
 BallDirx = 1
 'frmField.BackColor = vbMagenta
 txtDown.Visible = True
 imgP2.Visible = False
 Cap = 1050    'used for imgCap.Height
 Bot = imgBot.Height + picRit.Height - 100
 Div = 80     'for dividing ballx bally
 Strok = 1     'for incr. speed after hit
 Tim = 0    ' control tim.interval
 Vin1 = 0   'count no of rounds win
 Vin2 = 0
 timBall.Interval = 150
 
 imgP2.Height = 1200
imgP2.Width = 170
imgP1.Height = 1200
imgP1.Width = 170
imgBot.Top = frmField.Height - imgBot.Height - picRit.Height
picRit.Width = frmField.Width
 picRit.Top = frmField.Height - picRit.Height
 cmdCredit.Top = imgBot.Top + 70
 cmdCredit.Left = imgBot.Left + 100
 End Sub

Private Sub imgRod_Click()
End Sub

Private Sub timBall_Timer()
BallX = BallX + BallDirx * frmField.ScaleWidth / Div
If BallX < 0 Then     'ball will appear from right
  Beep
  BallX = frmField.ScaleWidth - shpBall.Width
  BallDirx = -1         'go leftward
  Sp2 = Sp2 + 1
'if ball hits from left to p2 & p1
ElseIf BallDirx = 1 Then
     If shpBall.Left > (imgP2.Left - shpBall.Width) Then
       If shpBall.Left < imgP2.Left Then
          If shpBall.Top > (imgP2.Top - shpBall.Height) Then
             If shpBall.Top < (imgP2.Top + imgP2.Height) Then
             Beep
             BallX = imgP2.Left - shpBall.Width
             BallDirx = -1   'go to leftward
             Strok = Strok + 1
             Tim = 1
            End If
          End If
      ElseIf BallX > frmField.ScaleWidth - shpBall.Width Then 'shpBall.Left > imgP2.Left Then
       BallX = 0
       BallDirx = 1    'go rightward appear from left
       Beep
       Sp1 = Sp1 + 1
     End If
   End If
   If shpBall.Left < (imgP1.Left) Then
         If shpBall.Left > (imgP1.Left - shpBall.Width) Then
           If shpBall.Top > (imgP1.Top - shpBall.Height) Then
             If shpBall.Top < (imgP1.Top + imgP1.Height) Then
             Beep
             BallX = imgP1.Left - shpBall.Width
             BallDirx = -1   'go to leftward
             Strok = Strok + 1
             Tim = 1
             End If
          End If
        ElseIf BallX > frmField.ScaleWidth - shpBall.Width Then 'shpBall.Left > imgP2.Left Then
        BallX = 0
        BallDirx = 1  'right appear from left
        Sp1 = Sp1 + 1
        Beep
     End If
    End If

   
'if ball hits from right to p2 & p1
ElseIf BallDirx = -1 Then
     If shpBall.Left > (imgP2.Left + imgP2.Width) Then
        If shpBall.Left < (imgP2.Left + imgP2.Width + shpBall.Width) Then 'frmField.ScaleWidth Then
           If shpBall.Top > (imgP2.Top - shpBall.Height) Then
             If shpBall.Top < (imgP2.Top + imgP2.Height) Then
             Beep
             BallX = imgP2.Left + imgP2.Width
             BallDirx = 1   'go to rightward
             Strok = Strok + 1
              Tim = 1
             End If
          End If
        ElseIf BallX > frmField.ScaleWidth - shpBall.Width Then 'shpBall.Left > imgP2.Left Then
        BallX = 0
        BallDirx = 1  'right
        Beep
        Sp1 = Sp1 + 1
     End If
    End If
If shpBall.Left > (imgP1.Left + imgP1.Width) Then
       If shpBall.Left < (imgP1.Left + imgP1.Width + shpBall.Width) Then
          If shpBall.Top > (imgP1.Top - shpBall.Height) Then
             If shpBall.Top < (imgP1.Top + imgP1.Height) Then
             Beep
             BallX = imgP1.Left + imgP1.Width
             BallDirx = 1   'go to leftward
             Strok = Strok + 1
              Tim = 1
            End If
          End If
      ElseIf BallX > frmField.ScaleWidth - shpBall.Width Then
       BallX = 0
       BallDirx = 1
       Sp1 = Sp1 + 1
       Beep
     End If
   End If


End If
shpBall.Left = BallX
'controlling vertical direction of ball
BallY = BallY + BallDiry * (frmField.ScaleHeight - Cap - Bot) / Div
If BallY < Cap Then
  Beep
  BallY = Cap
  BallDiry = 1
ElseIf BallY > (frmField.ScaleHeight - shpBall.Height - Bot) Then
Beep
BallY = frmField.ScaleHeight - shpBall.Height - Bot
BallDiry = -1

End If
If Tim = 1 Then      'enter only ones when strike
If Strok = 3 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 4 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 5 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 7 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 9 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 11 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 13 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 15 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 17 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 19 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 21 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 23 Then
timBall.Interval = timBall.Interval - 10
ElseIf Strok = 25 Then
timBall.Interval = timBall.Interval - 5
ElseIf Strok = 28 Then
timBall.Interval = timBall.Interval - 5
ElseIf Strok = 31 Then
timBall.Interval = timBall.Interval - 5
ElseIf Strok = 33 Then
timBall.Interval = timBall.Interval - 5
ElseIf Strok = 35 Then
timBall.Interval = timBall.Interval - 5
ElseIf Strok = 37 Then
timBall.Interval = timBall.Interval - 5
ElseIf Strok = 39 Then
timBall.Interval = timBall.Interval - 5
End If
Tim = 0     'becomes 1 when strike
End If
If timBall.Interval < 20 Then
timBall.Interval = 20
End If
shpBall.Top = BallY
'score
txtP1.Text = Str(Sp1)
txtP2.Text = Str(Sp2)

'change Round
If Sp1 = 25 Then
 Vin1 = Vin1 + 1
 If Vin1 = 1 Then
   'timBall.Interval = 0
  imgVin.Visible = True
  imgResult.Visible = True
  imgNull2.Top = imgVin.Top + imgVin.Height
  imgNull2.Left = imgResult.Left + imgResult.Width
  imgOne1.Top = imgVin.Top + imgVin.Height
  imgOne1.Left = imgResult.Left - imgOne1.Width
  txtDown.Top = imgOne1.Top + imgOne1.Height
  txtDown.Text = "START SECOND ROUND"
  shpBall.Visible = False
  imgP1.Visible = False
  imgP2.Visible = False
  timBall.Interval = 0
 imgOne1.Visible = True
 imgNull1.Visible = False
 'Sp1 = 0
 'Sp2 = 0
 'Strok = 0
 'timBall.Interval = 100
 
 'imgOne.Visible = False
 'imgTwo.Visible = True
 'frmField.BackColor = &HFFC0C0
 ElseIf Vin1 = 2 Then
  txtDown.Text = "START THIRD ROUND"
  shpBall.Visible = False
  imgP1.Visible = False
  imgP2.Visible = False
  timBall.Interval = 0
  imgVin.Visible = True
  imgResult.Visible = True
  txtDown.Top = imgOne1.Top + imgOne1.Height
  imgTwo1.Visible = True
  imgOne1.Visible = False
  imgTwo1.Top = imgVin.Top + imgVin.Height
  imgTwo1.Left = imgResult.Left - imgTwo1.Width
   If Vin2 = 0 Then
    imgNull2.Top = imgVin.Top + imgVin.Height
    imgNull2.Left = imgResult.Left + imgResult.Width
    
   ElseIf Vin2 = 1 Then
    imgOne2.Top = imgVin.Top + imgVin.Height
    imgOne2.Left = imgResult.Left + imgResult.Width
   End If

 End If
ElseIf Sp2 = 25 Then
 Vin2 = Vin2 + 1
  If Vin2 = 1 Then
    imgVin.Visible = True
    imgResult.Visible = True
    imgOne2.Top = imgVin.Top + imgVin.Height
    imgOne2.Left = imgResult.Left + imgResult.Width
    imgNull1.Top = imgOne2.Top
    imgNull1.Left = imgResult.Left - imgNull1.Width
    txtDown.Top = imgNull1.Top + imgNull1.Height
    txtDown.Text = "START SECOND ROUND"
    shpBall.Visible = False
    imgP1.Visible = False
    imgP2.Visible = False
    timBall.Interval = 0
    imgOne2.Visible = True
    imgNull2.Visible = False
    'imgOne.Visible = False
    'imgTwo.Visible = True
  ElseIf Vin2 = 2 Then
    txtDown.Top = imgNull1.Top + imgNull1.Height
    txtDown.Text = "START THIRD ROUND"
    shpBall.Visible = False
    imgP1.Visible = False
    imgP2.Visible = False
    timBall.Interval = 0
    imgVin.Visible = True
    imgResult.Visible = True
    imgTwo2.Top = imgVin.Top + imgVin.Height
    imgTwo2.Left = imgResult.Left + imgResult.Width
    imgTwo2.Visible = True
    imgOne2.Visible = False
    If Vin1 = 0 Then
      imgNull1.Top = imgTwo2.Top
      imgNull1.Left = imgResult.Left - imgNull1.Width
    ElseIf Vin1 = 1 Then
       imgOne1.Top = imgTwo2.Top
       imgOne1.Left = imgResult.Left - imgOne1.Width
    End If
  End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < frmField.ScaleHeight - imgP2.Height Then
imgP2.Top = Y - 200
End If
If imgP2.Top > (frmField.ScaleHeight - Bot - imgP2.Height) Then
imgP2.Top = (frmField.ScaleHeight - Bot - imgP2.Height)
End If
End Sub

Private Sub txtDown_Click()
'frmField.BackColor = vbBlue
If Vin1 = 1 Then
  imgOne1.Visible = True
  imgNull1.Visible = False
  imgOne1.Left = 3840
  imgOne1.Top = 120
  imgNull2.Left = 7920
  imgNull2.Top = 120
  Sp1 = 0
  Sp2 = 0
  Strok = 0
  timBall.Interval = 100
  frmField.BackColor = &HFFC0C0
  imgOne.Visible = False
  imgTwo.Visible = True
  txtDown.Top = cmdCredit.Top
  shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False
ElseIf Vin1 = 2 Then
  imgTwo1.Visible = True
  imgOne1.Visible = False
  imgTwo1.Left = 3840
  imgTwo1.Top = 120
  If Vin2 = 0 Then
   imgNull2.Left = 7920
   imgNull2.Top = 120
  ElseIf Vin2 = 1 Then
   imgOne2.Left = 7920
   imgOne2.Top = 120
   End If
  Sp1 = 0
  Sp2 = 0
  Strok = 0
  timBall.Interval = 70
  frmField.BackColor = &HFF8080
  imgTwo.Visible = False
  imgThree.Visible = True
  txtDown.Top = cmdCredit.Top
  shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False


ElseIf Vin2 = 1 Then
imgOne2.Visible = True
imgNull2.Visible = False
imgOne2.Left = 3840
imgOne2.Top = 120
imgNull1.Left = 7920
imgNull1.Top = 120
Sp1 = 0
Sp2 = 0
Strok = 0
timBall.Interval = 100
frmField.BackColor = &HFFC0C0
imgOne.Visible = False
imgTwo.Visible = True
txtDown.Top = cmdCredit.Top
shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False
ElseIf Vin2 = 2 Then
  imgTwo2.Visible = True
  imgOne2.Visible = False
  imgTwo2.Left = 3840
  imgTwo2.Top = 120
  If Vin1 = 0 Then
    imgNull1.Left = 7920
    imgNull1.Top = 120
   ElseIf Vin2 = 1 Then
    imgOne1.Left = 7920
    imgOne1.Top = 120
   End If
 Sp1 = 0
 Sp2 = 0
 Strok = 0
 timBall.Interval = 70
 frmField.BackColor = &HC0C0C08
 imgTwo.Visible = False
 imgThree.Visible = True
 txtDown.Top = cmdCredit.Top
 shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False

Else
timBall.Enabled = True
shpBall.Visible = True
imgP2.Visible = True
imgP1.Visible = True
txtDown.Top = picRit.Top - txtDown.Height
txtDown.Left = 5500
'frmField.Width = 12000
'frmField.Height = 7000
 End If
End Sub

Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
txtDown.Text = "PAUSE" 'Str(KeyCode)

If KeyCode = 38 Then       'up
 If imgP2.Top > Cap Then
 imgP2.Top = imgP2.Top - 400
 Else
 imgP2.Top = Cap
 End If
ElseIf KeyCode = 40 Then   'down
 If imgP2.Top < (frmField.ScaleHeight - imgP2.Height - Bot) Then
 imgP2.Top = imgP2.Top + 400
 Else
 imgP2.Top = (frmField.ScaleHeight - imgP2.Height - Bot)
 End If
End If

If KeyCode = 87 Then
 If imgP1.Top > Cap Then
 imgP1.Top = imgP1.Top - 400
 Else
 imgP1.Top = Cap
 End If
ElseIf KeyCode = 83 Then
 If imgP1.Top < (frmField.ScaleHeight - imgP1.Height - Bot) Then
 imgP1.Top = imgP1.Top + 400
 Else
 imgP1.Top = (frmField.ScaleHeight - imgP1.Height - Bot)
 End If
End If
End Sub

