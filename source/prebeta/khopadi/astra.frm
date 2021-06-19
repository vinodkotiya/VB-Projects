VERSION 5.00
Begin VB.Form frmastra 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BATTLE(DANGER GARDEN)"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   FillStyle       =   7  'Diagonal Cross
   Icon            =   "astra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "astra.frx":1CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer timKhopadi 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3120
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "Speed_"
      Height          =   495
      Left            =   3240
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
      Interval        =   100
      Left            =   1680
      Top             =   3120
   End
   Begin VB.Image imgCredit 
      Height          =   525
      Left            =   10680
      Picture         =   "astra.frx":3994
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image imgDg2 
      Height          =   720
      Left            =   240
      Picture         =   "astra.frx":5AA6
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDg1 
      Height          =   720
      Left            =   240
      Picture         =   "astra.frx":7770
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgKh1 
      Height          =   720
      Left            =   5400
      Picture         =   "astra.frx":943A
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgKh2 
      Height          =   720
      Left            =   5400
      Picture         =   "astra.frx":A304
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgGh1 
      Height          =   720
      Left            =   7440
      Picture         =   "astra.frx":B1CE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImgGh2 
      Height          =   720
      Left            =   7440
      Picture         =   "astra.frx":B4D8
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgBil1 
      Height          =   720
      Left            =   9480
      Picture         =   "astra.frx":B7E2
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgBil2 
      Height          =   720
      Left            =   9480
      Picture         =   "astra.frx":D4AC
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Tank1 
      Height          =   720
      Left            =   11040
      Picture         =   "astra.frx":F176
      Top             =   7680
      Width           =   720
   End
   Begin VB.Image Tank3 
      Height          =   720
      Left            =   11040
      Picture         =   "astra.frx":10E40
      Top             =   7680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Tank2 
      Height          =   720
      Left            =   11040
      Picture         =   "astra.frx":12B0A
      Top             =   7680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astras2 
      Height          =   720
      Left            =   3000
      Picture         =   "astra.frx":147D4
      Top             =   1920
      Width           =   720
   End
   Begin VB.Image Astras1 
      Height          =   720
      Left            =   2760
      Picture         =   "astra.frx":1649E
      Top             =   1800
      Width           =   720
   End
   Begin VB.Image Astran2 
      Height          =   720
      Left            =   2880
      Picture         =   "astra.frx":18168
      Top             =   360
      Width           =   720
   End
   Begin VB.Image Astran1 
      Height          =   720
      Left            =   2640
      Picture         =   "astra.frx":19E32
      Top             =   360
      Width           =   720
   End
   Begin VB.Image Astrase2 
      Height          =   720
      Left            =   4200
      Picture         =   "astra.frx":1BAFC
      Top             =   2160
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astrase1 
      Height          =   720
      Left            =   3360
      Picture         =   "astra.frx":1D7C6
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image Astrasw2 
      Height          =   720
      Left            =   2040
      Picture         =   "astra.frx":1F490
      Top             =   1920
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astrasw1 
      Height          =   720
      Left            =   1440
      Picture         =   "astra.frx":2115A
      Top             =   1920
      Width           =   720
   End
   Begin VB.Image Astrane2 
      Height          =   720
      Left            =   3840
      Picture         =   "astra.frx":22E24
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astrane1 
      Height          =   720
      Left            =   3000
      Picture         =   "astra.frx":24AEE
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Astranw2 
      Height          =   720
      Left            =   1440
      Picture         =   "astra.frx":267B8
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astranw1 
      Height          =   720
      Left            =   2040
      Picture         =   "astra.frx":28482
      Top             =   600
      Width           =   720
   End
   Begin VB.Image imgBoom3 
      Height          =   720
      Left            =   2760
      Picture         =   "astra.frx":2A14C
      Top             =   2880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astraw1 
      Height          =   720
      Left            =   2040
      Picture         =   "astra.frx":2BE16
      Top             =   1320
      Width           =   720
   End
   Begin VB.Image Astraw2 
      Height          =   720
      Left            =   2040
      Picture         =   "astra.frx":2DAE0
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image Astrae2 
      Height          =   720
      Left            =   3000
      Picture         =   "astra.frx":2F7AA
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Astrae1 
      Height          =   720
      Left            =   3000
      Picture         =   "astra.frx":31474
      Top             =   1320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBoom2 
      Height          =   720
      Left            =   6120
      Picture         =   "astra.frx":3313E
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBoom1 
      Height          =   720
      Left            =   4680
      Picture         =   "astra.frx":34E08
      Top             =   2880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgVin 
      Height          =   6600
      Left            =   1320
      Picture         =   "astra.frx":36AD2
      Top             =   840
      Visible         =   0   'False
      Width           =   10095
   End
   Begin VB.Image imgGhost 
      Height          =   6900
      Left            =   2760
      Picture         =   "astra.frx":41F36
      Top             =   1440
      Width           =   9000
   End
End
Attribute VB_Name = "frmastra"
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
Dim Bg As Single        'display Bg
Dim Lop As Integer, Bop As Integer

Private Sub cmdMin_Click()
If timBall.Interval < 400 Then
timBall.Interval = timBall.Interval + 10
Else
timBall.Interval = 400
End If
End Sub

Private Sub cmdPlus_Click()
If timBall.Interval > 5 Then
timBall.Interval = timBall.Interval - 10
Else
timBall.Interval = 10
End If
End Sub


Private Sub Form_Load()
 BallX = frmastra.ScaleWidth - Astrae1.Width
 BallY = frmastra.ScaleHeight - Astrae1.Height
 Astrae1.Left = BallX
 Astrae1.Top = BallY
 BallDiry = -1
 BallDirx = -1
 'frmAstra.BackColor = vbMagenta
 Temp = 70
 Khopadi2 = 0
 Lop = 0
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
'timKhopadi.Enabled = False
imgBoom2.Visible = False
'timBall.Enabled = True
imgBoom1.Visible = False   'will reset badi khopadi when mouse move
imgGhost.Visible = False
imgVin.Visible = False
imgBoom3.Visible = False
''for switching controls
ImgKh1.Visible = True
ImgKh2.Visible = False
ImgGh1.Visible = True
ImgGh2.Visible = False
imgBil2.Visible = True
ImgBil1.Visible = False
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
Unload frmBball
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
'imgVin.Visible = False
End Sub

Private Sub timBall_Timer()
Khopadi2 = Khopadi2 + 1
If Khopadi2 > 10 Then
 Khopadi2 = 0
 End If
Temp = timBall.Interval
BallX = BallX + BallDirx * frmastra.ScaleWidth / 50
If BallX < 0 Then
  Beep
  BallX = 0
  BallDirx = 1     'go rightward
   Lop = Lop + 1
ElseIf BallX > frmastra.ScaleWidth - Astrae1.Width Then
  Beep
  BallX = frmastra.ScaleWidth - Astrae1.Width
  BallDirx = -1
  Lop = Lop + 1
End If

BallY = BallY + BallDiry * (frmastra.ScaleHeight - 615) / 50    '615 is height of text bar
If BallY < cmdPlus.Top + cmdPlus.Height Then
  Lop = Lop + 1
  Beep
  BallY = 615
  BallDiry = 1
  If Lop > 80 Then
      timBall.Interval = 0
      Lop = 0   'display blast
      Bop = 0 'launch missile
      imgBoom1.Visible = True
      imgBoom1.Left = BallX
      imgBoom1.Top = BallY
      Astrae1.Visible = False
      imgBoom2.Visible = False
      Astrae2.Visible = False
      timKhopadi.Enabled = True
      timBall.Enabled = False
      Bg = Rnd(2)
      If Bg < 0.5 Then
       imgVin.Visible = True
      Else
      imgGhost.Visible = True
      End If
  End If
ElseIf BallY > (frmastra.ScaleHeight) - Astrae1.Height Then
  Beep
  BallY = (frmastra.ScaleHeight) - Astrae1.Height
  BallDiry = -1
  Lop = Lop + 1
End If

If Khopadi2 > 5 Then
    
   
  If Mx < BallX Then
   If My < BallY Then
   Astranw1.Visible = True
   Astranw1.Left = BallX
   Astranw1.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astrasw1.Visible = False
   Astrasw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   Astranw2.Visible = False
   ElseIf My > BallY Then
   Astrasw1.Visible = True
   Astrasw1.Left = BallX
   Astrasw1.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   Astrasw2.Visible = False
   Else
   Astraw1.Visible = True
   Astraw1.Left = BallX
   Astraw1.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   'Astrasw1.Visible = False
   Astrasw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   End If
  
  ElseIf Mx > BallX Then
  If My < BallY Then
    If Mx < (BallX + Astran1.Width) Then
     Astran1.Visible = True
     Astran1.Left = BallX
     Astran1.Top = BallY
     'Astran1.Visible = False
     Astran2.Visible = False
     Astras1.Visible = False
     Astras2.Visible = False
     Astraw1.Visible = False
     Astraw2.Visible = False
     Astrae1.Visible = False
     Astrae2.Visible = False
     Astrasw1.Visible = False
     Astrasw2.Visible = False
     Astrane1.Visible = False
     Astrane2.Visible = False
     Astrase1.Visible = False
     Astrase2.Visible = False
     Astranw1.Visible = False
     Astranw2.Visible = False
        
   Else
   Astrane1.Visible = True
   Astrane1.Left = BallX
   Astrane1.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   'Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   End If
  ElseIf My > BallY Then
     If Mx < (BallX + Astran1.Width) Then
     Astran1.Visible = True
     Astran1.Left = BallX
     Astran1.Top = BallY
     'Astran1.Visible = False
     Astran2.Visible = False
     Astras1.Visible = False
     Astras2.Visible = False
     Astraw1.Visible = False
     Astraw2.Visible = False
     Astrae1.Visible = False
     Astrae2.Visible = False
     Astrasw1.Visible = False
     Astrasw2.Visible = False
     Astrane1.Visible = False
     Astrane2.Visible = False
     Astrase1.Visible = False
     Astrase2.Visible = False
     Astranw1.Visible = False
     Astranw2.Visible = False
    Else
   Astrase1.Visible = True
   Astrase1.Left = BallX
   Astrase1.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astrase2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrasw1.Visible = False
   Astrasw2.Visible = False
    End If

   Else
   Astrae1.Visible = True
   Astrae1.Left = BallX
   Astrae1.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   'Astrae1.Visible = False
   Astrae2.Visible = False
   Astrasw1.Visible = False
   Astrasw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   End If
  End If
ElseIf Khopadi2 < 5 Then
  If Mx < BallX Then
   If My < BallY Then
   Astranw2.Visible = True
   Astranw2.Left = BallX
   Astranw2.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astranw1.Visible = False
   Astrasw1.Visible = False
   Astrasw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   ElseIf My > BallY Then
   Astrasw2.Visible = True
   Astrasw2.Left = BallX
   Astrasw2.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astrasw1.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   Else
   Astraw2.Visible = True
   Astraw2.Left = BallX
   Astraw2.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astrasw1.Visible = False
   'Astrasw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   End If
  
  ElseIf Mx > BallX Then
  If My < BallY Then
   Astrane2.Visible = True
   Astrane2.Left = BallX
   Astrane2.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrane1.Visible = False
   'Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   ElseIf My > BallY Then
   Astrase2.Visible = True
   Astrase2.Left = BallX
   Astrase2.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astraw1.Visible = False
   Astraw2.Visible = False
   Astrae1.Visible = False
   Astrae2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrase1.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrasw1.Visible = False
   Astrasw2.Visible = False
   Else
   Astrae2.Visible = True
   Astrae2.Left = BallX
   Astrae2.Top = BallY
   Astran1.Visible = False
   Astran2.Visible = False
   Astras1.Visible = False
   Astras2.Visible = False
   Astranw1.Visible = False
   Astranw2.Visible = False
   Astrae1.Visible = False
   'Astrae2.Visible = False
   Astrasw1.Visible = False
   Astrasw2.Visible = False
   Astrane1.Visible = False
   Astrane2.Visible = False
   Astrase1.Visible = False
   Astrase2.Visible = False
   End If
  End If
End If


''mouse aur khopadi barabar hone par rokne ke liye
If (My + 100) > BallY Then
 If (My + 100) < (BallY + 800) Then
    If (Mx + 100) > BallX Then
     If (Mx + 100) < (BallX + 800) Then
      timBall.Interval = 0
      Lop = 0   'display blast
      Bop = 0 'launch missile
      imgBoom1.Visible = True
      imgBoom1.Left = BallX
      imgBoom1.Top = BallY
      Astrae1.Visible = False
      imgBoom2.Visible = False
      Astrae2.Visible = False
      timKhopadi.Enabled = True
      timBall.Enabled = False
      Bg = Rnd(2)
      If Bg < 0.5 Then
       imgVin.Visible = True
      Else
      imgGhost.Visible = True
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

Bop = Bop + 1
Khopadi2 = Khopadi2 + 1
Astran1.Visible = False
Astran2.Visible = False
Astras1.Visible = False
Astras2.Visible = False
Astraw1.Visible = False
Astraw2.Visible = False
Astrae1.Visible = False
Astrae2.Visible = False
Astrasw1.Visible = False
Astrasw2.Visible = False
Astrane1.Visible = False
Astrane2.Visible = False
Astrase1.Visible = False
Astrase2.Visible = False
Astranw1.Visible = False
Astranw2.Visible = False

If Khopadi2 > 30 Then
 Khopadi2 = 0
 End If
If Lop > 34 Then
 imgBoom1.Visible = False
 imgBoom2.Visible = False
 imgBoom3.Visible = False
 
 If Bop > 50 Then
 Tank1.Visible = True
 Tank2.Visible = False
 Tank3.Visible = False
 BallX = Tank1.Left - 200
 BallY = Tank1.Top
 Astranw1.Top = BallY
 Astranw1.Left = BallX
 Astranw1.Visible = True
 
 
 BallDirx = -1
 BallDiry = -1
 timKhopadi.Enabled = False
 timBall.Enabled = True
 ElseIf Bop > 40 Then
 Tank2.Visible = True
 Tank1.Visible = False
 Tank3.Visible = False
 ElseIf Bop > 30 Then
 Tank3.Visible = True
 Tank2.Visible = False
 Tank1.Visible = False
 ElseIf Bop > 10 Then
 Tank2.Visible = True
 Tank1.Visible = False
 Tank3.Visible = False
 ElseIf Bop > 0 Then
 Tank1.Visible = True
 Tank2.Visible = False
 Tank3.Visible = False
 End If
End If
'display only twice
If Lop < 35 Then
If Khopadi2 > 20 Then
 imgBoom1.Visible = False
 imgBoom3.Visible = False
 imgBoom2.Left = BallX
 imgBoom2.Top = BallY
 imgBoom2.Visible = True
ElseIf Khopadi2 > 10 Then
 imgBoom1.Visible = False
 imgBoom2.Visible = False
 imgBoom3.Left = BallX
 imgBoom3.Top = BallY
 imgBoom3.Visible = True
ElseIf Khopadi2 > 0 Then
 
 imgBoom2.Visible = False
 imgBoom3.Visible = False
 imgBoom1.Left = BallX
 imgBoom1.Top = BallY
 imgBoom1.Visible = True
End If
Lop = Lop + 1
End If
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

