VERSION 5.00
Begin VB.Form frmBilli 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BILLI(DANGER GARDEN)"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   FillStyle       =   7  'Diagonal Cross
   Icon            =   "billi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "billi.frx":1CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   6150
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer timKhopadi 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   960
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "Speed_"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "Speed+"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer timBall 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   2640
   End
   Begin VB.Image imgCredit 
      Height          =   525
      Left            =   10680
      Picture         =   "billi.frx":3994
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image imgDg2 
      Height          =   720
      Left            =   240
      Picture         =   "billi.frx":5AA6
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDg1 
      Height          =   720
      Left            =   240
      Picture         =   "billi.frx":7770
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgKh1 
      Height          =   720
      Left            =   6000
      Picture         =   "billi.frx":943A
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgKh2 
      Height          =   720
      Left            =   6000
      Picture         =   "billi.frx":A304
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgGh1 
      Height          =   720
      Left            =   7800
      Picture         =   "billi.frx":B1CE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgGh2 
      Height          =   720
      Left            =   7800
      Picture         =   "billi.frx":B4D8
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgTank1 
      Height          =   720
      Left            =   9720
      Picture         =   "billi.frx":B7E2
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgTank2 
      Height          =   720
      Left            =   9720
      Picture         =   "billi.frx":D4AC
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgSleep3 
      Height          =   720
      Left            =   2880
      Picture         =   "billi.frx":F176
      Top             =   2400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgB1b 
      Height          =   720
      Left            =   4680
      Picture         =   "billi.frx":10E40
      Top             =   2520
      Width           =   720
   End
   Begin VB.Image imgB2b 
      Height          =   720
      Left            =   2400
      Picture         =   "billi.frx":12B0A
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image imgB2 
      Height          =   720
      Left            =   3000
      Picture         =   "billi.frx":147D4
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image imgB1 
      Height          =   720
      Left            =   5160
      Picture         =   "billi.frx":1649E
      Top             =   2520
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgSleep2 
      Height          =   720
      Left            =   6120
      Picture         =   "billi.frx":18168
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgSleep1 
      Height          =   720
      Left            =   3720
      Picture         =   "billi.frx":19E32
      Top             =   2400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgVin 
      Height          =   4305
      Left            =   600
      Picture         =   "billi.frx":1BAFC
      Top             =   1920
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Image imgGhost 
      Height          =   3990
      Left            =   6840
      Picture         =   "billi.frx":1F69B
      Top             =   2040
      Visible         =   0   'False
      Width           =   4050
   End
End
Attribute VB_Name = "frmBilli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BallX As Integer
Dim BallY As Integer
Dim BallDirx As Integer
Dim BallDiry As Integer
Dim Mx As Integer, My As Integer 'store X,Y
Dim Temp As Integer        'store timeinterval
Dim Khopadi2 As Integer    'maintain time for khopadi2
Dim Bg As Single, Bgpos As Integer       'display Bg
Private Sub cmdMin_Click()
If timBall.Interval < 400 Then
timBall.Interval = timBall.Interval + 5
Else
timBall.Interval = 400
End If
End Sub

Private Sub cmdPlus_Click()
If timBall.Interval > 5 Then
timBall.Interval = timBall.Interval - 5
Else
timBall.Interval = 10
End If
End Sub


Private Sub Form_Load()
 BallX = frmBilli.ScaleWidth - imgB1.Width
 BallY = frmBilli.ScaleHeight - imgB1.Height
 imgB1.Left = BallX
 imgB1.Top = BallY
 BallDiry = -1
 BallDirx = -1
 'frmBilli.BackColor = vbMagenta
 Temp = 70
 Khopadi2 = 0
 End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Mx = X
My = Y
If X < BallX Then
 BallDirx = -1
ElseIf X > BallX Then
 BallDirx = 1
 End If
If Y < BallY Then
 BallDiry = -1
ElseIf Y > BallY Then
 BallDiry = 1
 End If
timBall.Interval = Temp      'will restart the khopadi
timKhopadi.Enabled = False
imgSleep2.Visible = False
timBall.Enabled = True
imgSleep1.Visible = False   'will reset badi khopadi when mouse move
imgGhost.Visible = False
imgVin.Visible = False
imgSleep3.Visible = False
''for switch control
ImgKh1.Visible = True
ImgKh2.Visible = False
ImgGh1.Visible = True
ImgGh2.Visible = False
imgTank1.Visible = True
ImgTank2.Visible = False
imgDg2.Visible = False
imgDg1.Visible = True


End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload frmCredit
End Sub

Private Sub imgCredit_Click()
Load frmCredit
frmCredit.Visible = True

End Sub

Private Sub imgDg2_Click()
Load frmStart
frmStart.Visible = True
Unload Me
Unload frmBilli
Unload frmastra
Unload frmGhost
End Sub

Private Sub imgDg1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDg1.Visible = False
imgDg2.Visible = True
End Sub


Private Sub imgGhost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'imgGhost.Visible = False
imgVin.Visible = False
End Sub

Private Sub imgVin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGhost.Visible = False
imgVin.Visible = False
End Sub

Private Sub timBall_Timer()

Khopadi2 = Khopadi2 + 1
If Khopadi2 > 10 Then
 Khopadi2 = 0
 End If
Temp = timBall.Interval
BallX = BallX + BallDirx * frmBilli.ScaleWidth / 50
If BallX < 0 Then
  Beep
  BallX = 0
  BallDirx = 1     'go rightward
ElseIf BallX > frmBilli.ScaleWidth - imgB1.Width Then
  Beep
  BallX = frmBilli.ScaleWidth - imgB1.Width
  BallDirx = -1
End If
If Khopadi2 > 5 Then
  If Mx < BallX Then
  imgB2b.Visible = True
  imgB1b.Visible = False
  imgB1.Visible = False
  imgB2.Visible = False
  imgB2b.Left = BallX
  ElseIf Mx > BallX Then
  imgB2.Left = BallX
  imgB2.Visible = True
  imgB1.Visible = False
  imgB1b.Visible = False
  imgB2b.Visible = False
  End If
ElseIf Khopadi2 < 5 Then
 If Mx < BallX Then
  imgB1b.Visible = True
  imgB2b.Visible = False
  imgB1.Visible = False
  imgB2.Visible = False
  imgB1b.Left = BallX
 ElseIf Mx > BallX Then
 imgB1.Left = BallX
 imgB1.Visible = True
 imgB2.Visible = False
 imgB1b.Visible = False
 imgB2b.Visible = False
 End If
End If

 


BallY = BallY + BallDiry * (frmBilli.ScaleHeight - 615) / 50    '615 is height of text bar
If BallY < 615 Then
  Beep
  BallY = 615
  BallDiry = 1
ElseIf BallY > (frmBilli.ScaleHeight) - imgB1.Height Then
  Beep
  BallY = (frmBilli.ScaleHeight) - imgB1.Height
  BallDiry = -1
End If
If Khopadi2 > 5 Then
  If Mx < BallX Then
  imgB2b.Visible = True
  imgB1b.Visible = False
  imgB1.Visible = False
  imgB2.Visible = False
  imgB2b.Top = BallY
  ElseIf Mx > BallX Then
  imgB2.Top = BallY
  imgB2.Visible = True
  imgB1.Visible = False
  imgB1b.Visible = False
  imgB2b.Visible = False
  End If
ElseIf Khopadi2 < 5 Then
 If Mx < BallX Then
  imgB1b.Visible = True
  imgB2b.Visible = False
  imgB1.Visible = False
  imgB2.Visible = False
  imgB1b.Top = BallY
 ElseIf Mx > BallX Then
 imgB1.Top = BallY
 imgB1.Visible = True
 imgB2.Visible = False
 imgB1b.Visible = False
 imgB2b.Visible = False
 End If
End If



''mouse aur khopadi barabar hone par rokne ke liye
If (Mx + 100) > BallX Then
 If (Mx + 300) < (BallX + imgB1.Width) Then
    If (My + 100) > BallY Then
     If (My + 300) < (BallY + imgB1.Height) Then
      timBall.Interval = 0
      imgSleep2.Visible = True
      imgSleep2.Left = Mx
      imgSleep2.Top = My
      imgB1.Visible = False
      imgSleep1.Visible = False
      imgB2.Visible = False
      timKhopadi.Enabled = True
      timBall.Enabled = False
      Bg = Rnd(2)
      Bgpos = (Rnd(7000) * 10000)
      If Bg < 0.5 Then
       imgVin.Visible = True
       If Bgpos < (frmBilli.ScaleHeight - imgVin.Height) Then
        imgVin.Top = Bgpos
       End If
       If Bgpos < (frmBilli.ScaleWidth - imgVin.Width) Then
       imgVin.Left = Bgpos
       End If
      Else
      imgGhost.Visible = True
       If Bgpos < (frmBilli.ScaleHeight - imgGhost.Height) Then
       imgGhost.Top = Bgpos
       End If
       If Bgpos < (frmBilli.ScaleWidth - imgGhost.Width) Then
        imgGhost.Left = Bgpos
        End If
      End If
     Else
      timBall.Interval = Temp
     End If
    End If
  End If
End If
''Ek bar aur picha karne ke liye
If Mx < BallX Then
 BallDirx = -1
ElseIf Mx > BallX Then
 BallDirx = 1
 End If
If My < BallY Then
 BallDiry = -1
ElseIf My > BallY Then
 BallDiry = 1
 End If
End Sub

Private Sub timKhopadi_Timer()
Khopadi2 = Khopadi2 + 1
If Khopadi2 > 30 Then
 Khopadi2 = 0
 End If

imgB2.Visible = False
imgSleep1.Visible = False
imgB1b.Visible = False
imgB2b.Visible = False
If Khopadi2 > 20 Then
 imgB1.Visible = False
 imgSleep2.Visible = False
 imgSleep1.Visible = False
 imgSleep3.Left = BallX
 imgSleep3.Top = BallY
 imgSleep3.Visible = True
ElseIf Khopadi2 > 10 Then
 imgSleep1.Visible = False
 imgB1.Visible = False
 imgSleep3.Visible = False
 imgSleep2.Left = BallX
 imgSleep2.Top = BallY
 imgSleep2.Visible = True
ElseIf Khopadi2 > 0 Then
 
 imgSleep2.Visible = False
 imgB1.Visible = False
 imgSleep3.Visible = False
 imgSleep1.Left = BallX
 imgSleep1.Top = BallY
 imgSleep1.Visible = True
End If
End Sub




Private Sub imgTank1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgTank2.Visible = True
imgTank1.Visible = False
End Sub
Private Sub ImgTank2_Click()
Load frmastra
frmastra.Visible = True
Unload frmStart
Unload frmBball
Unload frmGhost
Unload frmBilli
End Sub



Private Sub ImgGh1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgGh2.Visible = True
ImgGh1.Visible = False
End Sub

Private Sub ImgGh2_Click()
Load frmGhost
frmGhost.Visible = True
Unload frmStart
Unload frmBball
Unload frmBilli
Unload frmastra
End Sub

Private Sub ImgKh2_Click()
Load frmBball
frmBball.Visible = True
Unload frmStart
Unload frmGhost
Unload frmBilli
Unload frmastra
End Sub

Private Sub ImgKh1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgKh2.Visible = True
ImgKh1.Visible = False
End Sub



