VERSION 5.00
Begin VB.Form frmField 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKey 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Text            =   "use arrow keys"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer timRace 
      Interval        =   90
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Image imgCar 
      Height          =   720
      Left            =   2640
      Picture         =   "frmField.frx":0000
      Top             =   6960
      Width           =   720
   End
   Begin VB.Image imgCar1 
      Height          =   720
      Left            =   1320
      Picture         =   "frmField.frx":0D42
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image imgRoad1 
      Height          =   9000
      Left            =   720
      Picture         =   "frmField.frx":1439
      Top             =   840
      Width           =   3450
   End
   Begin VB.Image imgRoad2 
      Height          =   9000
      Left            =   720
      Picture         =   "frmField.frx":3185
      Top             =   2400
      Width           =   3450
   End
   Begin VB.Image imgRoad 
      Height          =   9000
      Left            =   720
      Picture         =   "frmField.frx":4ED1
      Top             =   0
      Width           =   3450
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dir As Integer, CarY As Integer, DivY As Integer
Dim Rndm As Single
Dim Speed As Integer, Break As Integer
'Rndm = will show car1,2....Dir=downwards
'Break = if downkey press & again up then speed of car1,2 initial
'Speed =



Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
'txtKey.Text = "VIN"
If KeyCode = 37 Then   'left
  imgCar.Left = imgCar.Left - 60
  If KeyCode = 38 Then
  imgCar.Left = imgCar.Left - 110
  End If
ElseIf KeyCode = 39 Then
  imgCar.Left = imgCar.Left + 60
  If KeyCode = 38 Then
  imgCar.Left = imgCar.Left + 110
  
  End If
End If

If KeyCode = 38 Then    'up
  If Break = 1 Then
    Speed = 30
    Break = 0
    DivY = 200
    timRace.Enabled = True
    imgCar1.Visible = True
  End If
 Speed = Speed + 5
 DivY = DivY - 5
 Dir = 1
 ElseIf Speed > 0 Then
 Speed = Speed - 5
  DivY = DivY + 20
  Dir = -1
  
End If

If KeyCode = 40 Then 'break
  Speed = Speed - 5
  DivY = DivY - 5
  Dir = -1     'car1 go upward
  If imgCar1.Top < 50 Then
    imgCar1.Visible = False
    timRace.Enabled = False
  End If
'CarY = 0
Break = 1
End If
End Sub

Private Sub Form_Load()
Dir = 1
CarY = 100
Speed = 30
DivY = 200   'if more vel will less
End Sub


Private Sub timRace_Timer()
If DivY < 10 Then
 DivY = 20
End If
CarY = CarY + Dir * frmField.ScaleHeight / DivY
If CarY < 0 Then
  'Beep
  'CarY = frmField.ScaleHeight
  'Dir = -1        'go downward
ElseIf CarY > frmField.ScaleHeight - imgCar1.Height Then
  Beep
  Rndm = Rnd(500)
  Rndm = Rndm * 1000
  imgCar1.Left = imgRoad1.Left + Rndm
  'txtKey.Text = Str(Rndm)
  CarY = 0
  Dir = 1
  'BallY = picBall.ScaleHeight - shpBall.Height
  'Dir = -1
End If
imgCar1.Top = CarY
If imgRoad1.Top > frmField.ScaleHeight Then
  imgRoad1.Top = 0
ElseIf imgRoad1.Top < 0 Then
 imgRoad1.Top = frmField.ScaleHeight
Else
 imgRoad2.Top = imgRoad1.Top - imgRoad2.Height
End If
txtKey.Text = Str(Speed)
If Speed > 2000 Then
Speed = 2000
End If
imgRoad1.Top = imgRoad1.Top + Speed
imgRoad2.Top = imgRoad1.Top - imgRoad2.Height
txtKey.SetFocus
''car1 change the side automatically
Rndm = Rnd(1000)

  Rndm = Rndm * 100
  'txtKey.Text = Str(Rndm)
  If Rndm < 50 Then
   If imgCar1.Left < imgRoad1.Left Then
    imgCar1.Left = imgCar1.Left + 30
    Else
    imgCar1.Left = imgCar1.Left - 30
    End If
  Else
   If imgCar1.Left > (imgRoad1.Left + imgRoad1.Width - imgCar1.Width) Then
    imgCar1.Left = imgCar1.Left - 30
   Else
   imgCar1.Left = imgCar1.Left + 30
   End If
 End If
End Sub
