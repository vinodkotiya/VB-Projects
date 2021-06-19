VERSION 5.00
Begin VB.Form frmBball 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KHOPADI(DANGER GARDEN)"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   FillStyle       =   7  'Diagonal Cross
   Icon            =   "khopadi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "khopadi.frx":0ECA
   MousePointer    =   99  'Custom
   ScaleHeight     =   4275
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
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
      Interval        =   70
      Left            =   1680
      Top             =   2640
   End
   Begin VB.Image imgCredit 
      Height          =   525
      Left            =   10560
      Picture         =   "khopadi.frx":2B94
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image imgDg2 
      Height          =   720
      Left            =   240
      Picture         =   "khopadi.frx":4CA6
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDg1 
      Height          =   720
      Left            =   240
      Picture         =   "khopadi.frx":6970
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgGh1 
      Height          =   720
      Left            =   5040
      Picture         =   "khopadi.frx":863A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgGh2 
      Height          =   720
      Left            =   5040
      Picture         =   "khopadi.frx":8944
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgBil1 
      Height          =   720
      Left            =   6960
      Picture         =   "khopadi.frx":8C4E
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgBil2 
      Height          =   720
      Left            =   6960
      Picture         =   "khopadi.frx":A918
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgTank1 
      Height          =   720
      Left            =   8760
      Picture         =   "khopadi.frx":C5E2
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgTank2 
      Height          =   720
      Left            =   8760
      Picture         =   "khopadi.frx":E2AC
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgKhopadi2 
      Height          =   480
      Left            =   3000
      Picture         =   "khopadi.frx":FF76
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgKhopadi1 
      Height          =   480
      Left            =   5160
      Picture         =   "khopadi.frx":10280
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgKhopadir 
      Height          =   720
      Left            =   6000
      Picture         =   "khopadi.frx":1058A
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgKhopadi 
      Height          =   720
      Left            =   4080
      Picture         =   "khopadi.frx":11454
      Top             =   2400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgVin 
      Height          =   7905
      Left            =   1440
      Picture         =   "khopadi.frx":1231E
      Top             =   2280
      Visible         =   0   'False
      Width           =   8400
   End
   Begin VB.Image imgGhost 
      Height          =   6315
      Left            =   2880
      Picture         =   "khopadi.frx":257EA
      Top             =   2400
      Visible         =   0   'False
      Width           =   6675
   End
End
Attribute VB_Name = "frmBball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
'---------------------------------------------------------
'---------------- BY VINOD KOTIYA --------------------
'------ The code of this program is not very efficient
'------ because it was created on the early days of my
'------- visual basic computer programming.
'------- i made this programme without reading any VB book
'------- on the basis of my C++ experience i generally used
'------- if else statements
'------ code is easy and you can modify it
'------- in to a good code
'-------------------------------------------------------
'------ address S-2 shrimaya apartment sector-B/363
'------ sarvdharm colony bhopal
'---- fone +91-0755-2794428
'------ web: http://vinodkotiya.tripod.com     (without WWW)
'---- mail vinodkotiya24@rediffmail.com
'--------------------------------------------------------
'--------------------------------------------------------
Option Explicit
Dim BallX As Integer
Dim BallY As Integer
Dim BallDirx As Integer
Dim BallDiry As Integer
Dim Mx As Integer, My As Integer 'store X,Y
Dim Temp As Integer        'store timeinterval
Dim Khopadi2 As Integer    'maintain time for khopadi2
Dim Bg As Single, Bgpos As Single       'display Bg
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
 BallX = frmBball.ScaleWidth - imgKhopadi1.Width
 BallY = frmBball.ScaleHeight - imgKhopadi1.Height
 imgKhopadi1.Left = BallX
 imgKhopadi1.Top = BallY
 BallDiry = -1
 BallDirx = -1
 'frmBball.BackColor = vbMagenta
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
imgKhopadir.Visible = False
timBall.Enabled = True
imgKhopadi.Visible = False   'will reset badi khopadi when mouse move
imgGhost.Visible = False
imgVin.Visible = False
''for switching controls
ImgGh1.Visible = True
ImgGh2.Visible = False
imgBil2.Visible = True
ImgBil1.Visible = False
imgTank1.Visible = True
ImgTank2.Visible = False
imgDg1.Visible = True
imgDg2.Visible = False
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

'Mx = X
'My = Y
'If X < BallX Then
 'BallDirx = -1
'ElseIf X > BallX Then
' BallDirx = 1
 'End If
'If Y < BallY Then
 'BallDiry = -1
'ElseIf Y > BallY Then
 'BallDiry = 1
 'End If
'timBall.Interval = Temp      'will restart the khopadi
'timKhopadi.Enabled = False
'imgKhopadir.Visible = False
'timBall.Enabled = True
'imgKhopadi.Visible = False   'will reset badi khopadi when mouse move
imgGhost.Visible = False
imgVin.Visible = False
End Sub

Private Sub imgVin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgGhost.Visible = False
imgVin.Visible = False
End Sub

Private Sub timBall_Timer()
Khopadi2 = Khopadi2 + 1
If Khopadi2 > 19 Then
 Khopadi2 = 0
 End If
Temp = timBall.Interval
BallX = BallX + BallDirx * frmBball.ScaleWidth / 50
If BallX < 0 Then
  Beep
  BallX = 0
  BallDirx = 1     'go rightward
ElseIf BallX > frmBball.ScaleWidth - imgKhopadi1.Width Then
  Beep
  BallX = frmBball.ScaleWidth - imgKhopadi1.Width
  BallDirx = -1
End If
If Khopadi2 > 10 Then
 imgKhopadi2.Left = BallX
 imgKhopadi1.Visible = False
 imgKhopadi2.Visible = True
 ElseIf Khopadi2 < 10 Then
 imgKhopadi1.Left = BallX
 imgKhopadi2.Visible = False
 imgKhopadi1.Visible = True
 
End If

BallY = BallY + BallDiry * (frmBball.ScaleHeight - 615) / 50    '615 is height of text bar
If BallY < 615 Then
  Beep
  BallY = 615
  BallDiry = 1
ElseIf BallY > (frmBball.ScaleHeight) - imgKhopadi1.Height Then
  Beep
  BallY = (frmBball.ScaleHeight) - imgKhopadi1.Height
  BallDiry = -1
End If
If Khopadi2 > 10 Then
 imgKhopadi2.Top = BallY
 imgKhopadi2.Visible = True
 imgKhopadi1.Visible = False
ElseIf Khopadi2 < 10 Then
 imgKhopadi1.Top = BallY
 imgKhopadi1.Visible = True
 imgKhopadi2.Visible = False
End If



''mouse aur khopadi barabar hone par rokne ke liye
If (Mx + 100) > BallX Then
 If (Mx + 100) < (BallX + imgKhopadi1.Width) Then
    If (My + 100) > BallY Then
     If (My + 100) < (BallY + imgKhopadi1.Height) Then
      timBall.Interval = 0
      imgKhopadir.Visible = True
      imgKhopadir.Left = BallX
      imgKhopadir.Top = BallY
      imgKhopadi1.Visible = False
      imgKhopadi.Visible = False
      imgKhopadi2.Visible = False
      timKhopadi.Enabled = True
      timBall.Enabled = False
      Bg = Rnd(2)
      Bgpos = (Rnd(7000) * 10000)
      If Bg < 0.5 Then
       imgVin.Visible = True
       If Bgpos < (frmBball.ScaleHeight - imgVin.Height) Then
        imgVin.Top = Bgpos
       End If
       If Bgpos < (frmBball.ScaleWidth - imgVin.Width) Then
       imgVin.Left = Bgpos
       End If
      Else
      imgGhost.Visible = True
       If Bgpos < (frmBball.ScaleHeight - imgGhost.Height) Then
       imgGhost.Top = Bgpos
       End If
       If Bgpos < (frmBball.ScaleWidth - imgGhost.Width) Then
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
If Khopadi2 = 40 Then
 Khopadi2 = 0
 End If
If Khopadi2 > 30 Then
 
 imgKhopadir.Visible = False
 imgKhopadi2.Visible = False
 imgKhopadi.Visible = False
 imgKhopadi1.Left = BallX
 imgKhopadi1.Top = BallY
 imgKhopadi1.Visible = True
ElseIf Khopadi2 > 20 Then
 imgKhopadi1.Visible = False
 imgKhopadir.Visible = False
 imgKhopadi.Visible = False
 imgKhopadi2.Left = BallX
 imgKhopadi2.Top = BallY
 imgKhopadi2.Visible = True
ElseIf Khopadi2 > 10 Then
 imgKhopadi.Visible = False
 imgKhopadi1.Visible = False
 imgKhopadi2.Visible = False
 imgKhopadir.Left = BallX
 imgKhopadir.Top = BallY
 imgKhopadir.Visible = True
ElseIf Khopadi2 > 0 Then
 
 imgKhopadir.Visible = False
 imgKhopadi1.Visible = False
 imgKhopadi2.Visible = False
 imgKhopadi.Left = BallX
 imgKhopadi.Top = BallY
 imgKhopadi.Visible = True
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
Private Sub ImgBil1_Click()
Load frmBilli
frmBilli.Visible = True
Unload frmStart
Unload frmBball
Unload frmGhost
Unload frmastra
End Sub


Private Sub imgBil2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgBil1.Visible = True
imgBil2.Visible = False
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

