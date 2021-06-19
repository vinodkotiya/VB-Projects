VERSION 5.00
Begin VB.Form frmField 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VINTRAP (DEMO VER.)"
   ClientHeight    =   8130
   ClientLeft      =   2085
   ClientTop       =   930
   ClientWidth     =   11880
   Icon            =   "vintrap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "vintrap.frx":0ECA
   MousePointer    =   99  'Custom
   Picture         =   "vintrap.frx":101C
   ScaleHeight     =   8130
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timFun 
      Interval        =   50
      Left            =   1320
      Top             =   3360
   End
   Begin VB.TextBox txtDown 
      BackColor       =   &H00000000&
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
      Height          =   495
      Left            =   -120
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Text            =   "if controls of player1 are not functoning properly,press Tab"
      Top             =   3960
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.TextBox txtVinner 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   480
      Left            =   4320
      TabIndex        =   6
      Text            =   "VINOD"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtNm2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4320
      MouseIcon       =   "vintrap.frx":245E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtNm1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4300
      MouseIcon       =   "vintrap.frx":2768
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtStart 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Text            =   "START"
      Top             =   4560
      Visible         =   0   'False
      Width           =   4455
   End
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
      TabIndex        =   2
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
      TabIndex        =   1
      Text            =   "0"
      Top             =   615
      Width           =   615
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   7560
      Width           =   11895
   End
   Begin VB.Timer timBall 
      Enabled         =   0   'False
      Interval        =   170
      Left            =   600
      Top             =   3360
   End
   Begin VB.Image shpBall 
      Height          =   480
      Left            =   840
      Picture         =   "vintrap.frx":2A72
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image cmdCredit 
      Height          =   720
      Left            =   8160
      Picture         =   "vintrap.frx":2D7C
      Top             =   6430
      Width           =   720
   End
   Begin VB.Image imgButton2 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "vintrap.frx":3C46
      MousePointer    =   99  'Custom
      Picture         =   "vintrap.frx":3F50
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgButton1 
      Height          =   480
      Left            =   4680
      Picture         =   "vintrap.frx":4B92
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image imgAccelL 
      Height          =   480
      Left            =   5880
      Picture         =   "vintrap.frx":57D4
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAccelR 
      Height          =   480
      Left            =   5400
      Picture         =   "vintrap.frx":61C2
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   1785
      Left            =   3120
      MouseIcon       =   "vintrap.frx":6BB0
      MousePointer    =   99  'Custom
      Picture         =   "vintrap.frx":6EBA
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Image imgTwo2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "vintrap.frx":11108
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgTwo1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "vintrap.frx":11FDA
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgThree 
      Height          =   330
      Left            =   6960
      Picture         =   "vintrap.frx":12EAC
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgResult 
      Height          =   750
      Left            =   3840
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap.frx":13556
      Top             =   3000
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgVin 
      Height          =   750
      Left            =   480
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap.frx":17818
      Top             =   2160
      Visible         =   0   'False
      Width           =   10800
   End
   Begin VB.Image imgTwo 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   6960
      Picture         =   "vintrap.frx":208FA
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgOne2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "vintrap.frx":20FA4
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgOne1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "vintrap.frx":21E76
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgNull2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "vintrap.frx":22D48
      Top             =   120
      Width           =   840
   End
   Begin VB.Image imgNull1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "vintrap.frx":23C1A
      Top             =   120
      Width           =   840
   End
   Begin VB.Image imgBot 
      Height          =   1200
      Left            =   0
      MousePointer    =   3  'I-Beam
      Picture         =   "vintrap.frx":24AEC
      Top             =   6360
      Width           =   12000
   End
   Begin VB.Image imgOne 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   6960
      Picture         =   "vintrap.frx":3492E
      Top             =   360
      Width           =   435
   End
   Begin VB.Image imgCap 
      Height          =   1050
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap.frx":34FD8
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image imgP1 
      Height          =   1920
      Left            =   1440
      Picture         =   "vintrap.frx":42EDA
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgP2 
      Height          =   1920
      Left            =   10080
      Picture         =   "vintrap.frx":43B1C
      Top             =   4440
      Width           =   240
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
Dim Vin1 As Integer, Vin2 As Integer, Final As Integer
Dim Nm1 As String, Nm2 As String
Dim Temp As Integer, restorespeed As Integer, incrspeed As Integer
Dim Restart As Integer
Dim Xrnd As Single, Yrnd As Single     'store random no
Dim Hit As Integer, Race As Integer, Lace As Integer
Dim Com As Integer   'for opponent
Dim Level As Integer  ' For hardness

Private Sub cmdCredit_Click()
frmCredit.Visible = True
frmField.Visible = True
'timBall.Interval = 0
End Sub



Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub File1_Click()

End Sub

Private Sub Form_Load()
 Restart = 1
 BallX = frmField.ScaleWidth / 2
 BallY = frmField.ScaleHeight / 2
 shpBall.Left = BallX
 shpBall.Top = BallY
 shpBall.Height = 430
 shpBall.Width = 430
 BallDiry = 1
 BallDirx = 1
' frmField.BackColor = &HC0FFC0
imgAccelR.Height = 300 'reduce height of 32 x 32 icon
imgAccelL.Height = 300
 txtDown.Visible = True
 imgP2.Visible = False
 Cap = 1050    'used for imgCap.Height
 Bot = imgBot.Height + picScroll.Height - 100
 Div = 80     'for dividing ballx bally
 Strok = 1     'for incr. speed after hit
 Tim = 0    ' control tim.interval
 Vin1 = 0   'count no of rounds win
 Vin2 = 0
 'imgP2.Height = 1200
 'imgP2.Width = 170
 'imgP1.Height = 1200
 'imgP1.Width = 170
imgP1.Left = 1550
imgP2.Left = frmField.ScaleWidth - 1550 - imgP2.Width
imgBot.Top = frmField.Height - imgBot.Height - picScroll.Height
picScroll.Width = frmField.Width
picScroll.Top = frmField.Height - picScroll.Height
cmdCredit.Top = imgBot.Top + 70
cmdCredit.Left = 8160
 
 imgOne1.Visible = False
 imgTwo1.Visible = False
 imgOne2.Visible = False
 imgTwo2.Visible = False
 imgTwo.Visible = False
 imgThree.Visible = False
 imgVin.Visible = False
 imgResult.Visible = False
 Temp = 150
 'Com = 1
 End Sub

Private Sub imgRod_Click()
End Sub



Private Sub imgResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton1.Visible = True
imgButton2.Visible = False

End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton1.Visible = True
imgButton2.Visible = False
End Sub



Private Sub imgButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton2.Left = imgButton1.Left
imgButton2.Top = imgButton1.Top
imgButton1.Visible = False

imgButton2.Visible = True
End Sub



Private Sub timBall_Timer()
'txtDown.Text = " "
'Temp = txtDown_Click()
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
  BallDiry = 1 'go downward
ElseIf BallY > (frmField.ScaleHeight - shpBall.Height - Bot) Then
Beep
BallY = frmField.ScaleHeight - shpBall.Height - Bot
BallDiry = -1   'go upward

End If

If Tim = 1 Then      'enter only ones when strike
 If restorespeed = 1 Then
  timBall.Interval = Temp
  restorespeed = 0
  Hit = 0   'next accelarator will display
 End If
 'txtP1.Text = Str(timBall.Interval)
 If timBall.Interval > 20 Then
  If Strok Mod 2 = 0 Then     ' incr speed when strike
  timBall.Interval = timBall.Interval - 10
  End If
  Else
   timBall.Interval = 20
 End If
 
Tim = 0     'becomes 1 when strike
End If      'tim ended
If incrspeed = 1 Then
 Temp = timBall.Interval ' stores the current speed & make speed 10 when strike restore speed & next accelarator will display
 timBall.Interval = 10
 restorespeed = 1
 incrspeed = 0
' txtDown.Text = Str(Temp)
End If
shpBall.Top = BallY
'score
txtP1.Text = Str(Sp1)
txtP2.Text = Str(Sp2)

'change Round
If Sp1 = 25 Then
 Vin1 = Vin1 + 1
 imgP1.Height = imgP1.Height + 140 'bONUS INCR HIEGHT
 frmField.BackColor = &HC0FFFF
 txtVinner.Visible = True
 txtVinner.Text = Nm1
 shpBall.Visible = False
 imgP1.Visible = False
 imgP2.Visible = False
 timBall.Interval = 0
 imgVin.Visible = True
 imgResult.Visible = True
 txtDown.Top = imgResult.Top + imgResult.Height
 txtDown.Width = imgResult.Width + 1680
 txtDown.Left = imgResult.Left - 840
 imgButton1.Visible = True
 imgButton1.Top = txtDown.Top + txtDown.Height
 imgButton1.Left = imgResult.Left + txtDown.Width / 2
 

 If Vin1 = 1 Then
   'timBall.Interval = 0
  
  'txtDown.Top = imgResult.Top + imgResult.Height
  'txtDown.Width = 4700
  'txtDown.Left = imgResult.Left
  'txtDown.Text = "START SECOND ROUND"
  'shpBall.Visible = False
  'imgP1.Visible = False
  'imgP2.Visible = False
  'timBall.Interval = 0
  imgOne1.Visible = True
  imgNull1.Visible = False
  'imgVin.Visible = True
  'imgResult.Visible = True
  imgOne1.Top = imgVin.Top + imgVin.Height
  imgOne1.Left = imgResult.Left - imgOne1.Width
  If Vin2 = 0 Then
   imgNull2.Visible = True
   imgNull2.Top = imgVin.Top + imgVin.Height
   imgNull2.Left = imgResult.Left + imgResult.Width
   txtDown.Text = "START SECOND ROUND"
   ElseIf Vin2 = 1 Then
    imgOne2.Visible = True
    imgNull2.Visible = False
    imgOne2.Top = imgVin.Top + imgVin.Height
   imgOne2.Left = imgResult.Left + imgResult.Width
   txtDown.Text = "START FINAL ROUND"
   ElseIf Vin2 = 2 Then
   Final = 2
   End If
     'Sp1 = 0
 'Sp2 = 0
 'Strok = 0
 'timBall.Interval = 100
 
 'imgOne.Visible = False
 'imgTwo.Visible = True
 'frmField.BackColor = &HFFC0C0
 ElseIf Vin1 = 2 Then
  txtDown.Text = "CONGRATULATIONS " & Nm1 & " YOU ARE THE WINNER"
  Final = 1
  'shpBall.Visible = False
  'imgP1.Visible = False
  'imgP2.Visible = False
  'timBall.Interval = 0
  'imgVin.Visible = True
  'imgResult.Visible = True
 ' txtDown.Top = imgOne1.Top + imgOne1.Height
  imgTwo1.Visible = True
  imgOne1.Visible = False
  imgTwo1.Top = imgVin.Top + imgVin.Height
  imgTwo1.Left = imgResult.Left - imgTwo1.Width
   If Vin2 = 0 Then
    imgNull2.Visible = True
    imgNull2.Top = imgVin.Top + imgVin.Height
    imgNull2.Left = imgResult.Left + imgResult.Width
    Final = 1  'player 1 wins game
   ElseIf Vin2 = 1 Then
    imgOne2.Visible = True
    imgNull2.Visible = False
    imgOne2.Top = imgVin.Top + imgVin.Height
    imgOne2.Left = imgResult.Left + imgResult.Width
    
    Final = 1
   End If

 End If
ElseIf Sp2 = 25 Then
 Vin2 = Vin2 + 1
timBall.Interval = 0
 imgVin.Visible = True
 imgResult.Visible = True
 imgP2.Height = imgP2.Height + 140  'BONUS INCR HEIGHT
 frmField.BackColor = &HC0FFFF
 txtVinner.Visible = True
 txtVinner.Text = Nm2
 txtDown.Top = imgResult.Top + imgResult.Height
 txtDown.Width = imgResult.Width + 1680
 txtDown.Left = imgResult.Left - 840
 shpBall.Visible = False
 imgP1.Visible = False
 imgP2.Visible = False
 
  imgButton1.Visible = True
  imgButton1.Top = txtDown.Top + txtDown.Height
  imgButton1.Left = imgResult.Left + txtDown.Width / 2
  If Vin2 = 1 Then
    'txtDown.Top = imgResult.Top + imgResult.Height
    'txtDown.Width = 4700
    'txtDown.Left = imgResult.Left
    'shpBall.Visible = False
    'imgP1.Visible = False
    'imgP2.Visible = False
    'timBall.Interval = 0
    imgOne2.Visible = True
    imgNull2.Visible = False
    'imgVin.Visible = True
    'imgResult.Visible = True
    imgOne2.Top = imgVin.Top + imgVin.Height
    imgOne2.Left = imgResult.Left + imgResult.Width
    If Vin1 = 0 Then
     imgNull1.Visible = True
     imgNull1.Top = imgOne2.Top
     imgNull1.Left = imgResult.Left - imgNull1.Width
     txtDown.Text = "START SECOND ROUND"
    ElseIf Vin1 = 1 Then
     imgOne1.Visible = True
     imgNull1.Visible = False
     imgOne1.Top = imgOne2.Top
     imgOne1.Left = imgResult.Left - imgOne1.Width
     txtDown.Text = "START FINAL ROUND"
    ElseIf Vin1 = 2 Then
     Final = 1
     End If
    'imgOne.Visible = False
    'imgTwo.Visible = True
  ElseIf Vin2 = 2 Then
    'txtDown.Top = imgNull1.Top + imgNull1.Height
     txtDown.Text = "CONGRATULATIONS " & Nm2 & " YOU ARE THE WINNER"
    'shpBall.Visible = False
    'imgP1.Visible = False
    'imgP2.Visible = False
    'timBall.Interval = 0
    'imgVin.Visible = True
    'imgResult.Visible = True
    imgTwo2.Top = imgVin.Top + imgVin.Height
    imgTwo2.Left = imgResult.Left + imgResult.Width
    imgTwo2.Visible = True
    imgOne2.Visible = False
    Final = 2
    If Vin1 = 0 Then
      imgNull1.Visible = True
      imgNull1.Top = imgTwo2.Top
      imgNull1.Left = imgResult.Left - imgNull1.Width
      Final = 2 'player 2 wins the game
    ElseIf Vin1 = 1 Then
        imgOne1.Visible = True
        imgNull1.Visible = False
       imgOne1.Top = imgTwo2.Top
       imgOne1.Left = imgResult.Left - imgOne1.Width
    End If
  End If
End If
'txtDown.Text = "PAUSE" 'Str(KeyCode)
'Fun
'If Strok \ 6 = 0 Then
 
  If Hit = 0 Then 'ball has hit it
   If Race = 1 Then   'right accln will display
      Yrnd = Rnd(4353443)
     If (Yrnd * 10000) > (frmField.ScaleTop + Cap + imgAccelR.Height) Then 'shold display b/n field
       If (Yrnd * 10000) < (frmField.ScaleHeight - Bot - imgAccelR.Height) Then
       imgAccelR.Visible = True
       imgAccelR.Top = (Yrnd * 10000)
       Race = 0
       Hit = 1
       End If
     Else
     Hit = 0
     Race = 1
     End If
     End If
    If Lace = 1 Then
    Yrnd = Rnd(4353443)
     If (Yrnd * 10000) > (frmField.ScaleTop + Cap + imgAccelL.Height) Then 'shold display b/n field
       If (Yrnd * 10000) < (frmField.ScaleHeight - Bot - imgAccelL.Height) Then
        imgAccelL.Visible = True
        imgAccelL.Top = (Yrnd * 10000)
        Lace = 0
        Hit = 1
        End If
     Else
      Hit = 0
      Lace = 1
     End If
   End If
  End If
 'End If
 'ball come  to hit the accelaretors from 4 sides
If Hit = 1 Then
  If BallDiry = -1 Then   'upward
      If shpBall.Top > (imgAccelR.Top) Then
        If shpBall.Top < (imgAccelR.Top + imgAccelR.Height + shpBall.Height) Then
          If shpBall.Left > (imgAccelR.Left) Then
             If (shpBall.Left - shpBall.Width) < (imgAccelR.Left + imgAccelR.Width) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = 1
             BallDirx = 1  'only rightwards
             Lace = 1 'left accn will display
             Race = 0
             imgAccelR.Visible = False
             imgAccelR.Top = 0
             incrspeed = 1
             End If
          End If
         End If
       End If
     If shpBall.Top > (imgAccelL.Top) Then
        If shpBall.Top < (imgAccelL.Top + imgAccelL.Height + shpBall.Height) Then
          If shpBall.Left > (imgAccelL.Left) Then
             If (shpBall.Left - shpBall.Width) < (imgAccelL.Left + imgAccelL.Width) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = 1
             BallDirx = -1
             Race = 1 'rightaccn will display
             Lace = 0
             imgAccelL.Visible = False
             imgAccelL.Top = 0
             incrspeed = 1
             End If
          End If
         End If
       End If
   'b/c accel are short so no need of dirx only imp is diry
       'code is similar for both diry +-1
   If shpBall.Left > (imgAccelR.Left - shpBall.Width) Then
       If shpBall.Left < imgAccelR.Left Then
          If shpBall.Top > (imgAccelR.Top - shpBall.Height) Then
             If shpBall.Top < (imgAccelR.Top + imgAccelR.Height) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = -1
             BallDirx = 1
             Lace = 1 'leftaccn will display
             Race = 0
             imgAccelR.Visible = False
             imgAccelR.Top = 0
             incrspeed = 1
              End If
          End If
        End If
      End If
   If shpBall.Left > (imgAccelL.Left - shpBall.Width) Then
       If shpBall.Left < imgAccelL.Left Then
          If shpBall.Top > (imgAccelL.Top - shpBall.Height) Then
             If shpBall.Top < (imgAccelL.Top + imgAccelL.Height) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = -1
             BallDirx = -1
             Race = 1 'rightaccn will display
             Lace = 0
             imgAccelL.Visible = False
             imgAccelL.Top = 0
             incrspeed = 1
             End If
          End If
        End If
      End If

    End If  'balldiry = -1 ended
    
  If BallDiry = 1 Then 'Downward
   If shpBall.Top > (imgAccelR.Top - shpBall.Height) Then
        If shpBall.Top < (imgAccelR.Top) Then
          If shpBall.Left > (imgAccelR.Left) Then
             If (shpBall.Left - shpBall.Width) < (imgAccelR.Left + imgAccelR.Width) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = -1
             BallDirx = 1  'only rightwards
             Lace = 1 'left accn will display
             Race = 0
             imgAccelR.Visible = False
             imgAccelR.Top = 0
             incrspeed = 1
             End If
          End If
         End If
        End If
      If shpBall.Top > (imgAccelL.Top - shpBall.Height) Then
        If shpBall.Top < (imgAccelL.Top) Then
          If shpBall.Left > (imgAccelL.Left) Then
             If (shpBall.Left - shpBall.Width) < (imgAccelL.Left + imgAccelL.Width) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = -1
             BallDirx = -1
             Race = 1 'rightaccn will display
             Lace = 0
             imgAccelL.Visible = False
             imgAccelL.Top = 0
             incrspeed = 1
             End If
          End If
         End If
       End If
  'b/c accel are short so no need of dirx only imp is diry
       'code is similar for both diry +-1
   If shpBall.Left > (imgAccelR.Left - shpBall.Width) Then
       If shpBall.Left < imgAccelR.Left Then
          If shpBall.Top > (imgAccelR.Top - shpBall.Height) Then
             If shpBall.Top < (imgAccelR.Top + imgAccelR.Height) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = -1
             BallDirx = 1
             Race = 0 'lefttaccn will display
             Lace = 1
             imgAccelR.Visible = False
             imgAccelR.Top = 0
             incrspeed = 1
             End If
          End If
        End If
      End If
   If shpBall.Left > (imgAccelL.Left - shpBall.Width) Then
       If shpBall.Left < imgAccelL.Left Then
          If shpBall.Top > (imgAccelL.Top - shpBall.Height) Then
             If shpBall.Top < (imgAccelL.Top + imgAccelL.Height) Then
             Beep
             'Hit = 0   'ball has hit
             BallDiry = -1
             BallDirx = -1
             Race = 1 'rightaccn will display
             Lace = 0
             imgAccelL.Visible = False
             imgAccelL.Top = 0
             incrspeed = 1
             End If
          End If
        End If
      End If
        
     End If 'Ball Diry = 1 ended
  
Else
   'Hit = 0
   'Lace = 1
   'Race = 0
    End If    'hit ended
''computer play
If Com = 1 Then
  If BallDiry = -1 Then
   If imgP1.Top > Cap Then
    imgP1.Top = imgP1.Top - Level
   Else
    imgP1.Top = Cap
   End If
  ElseIf BallDiry = 1 Then
   If imgP1.Top < (frmField.ScaleHeight - imgP1.Height - Bot) Then
     imgP1.Top = imgP1.Top + Level
     Else
     imgP1.Top = (frmField.ScaleHeight - imgP1.Height - Bot)
    End If
  End If
End If     'End com
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
'txtDown.Top = cmdCredit.Top
'txtDown.Left = cmdCredit.Left
'txtDown.Width = cmdCredit.Width
txtDown.Top = picScroll.Top
txtDown.Left = picScroll.Left
'txtDown.Width = txtNm1.Width
'txtDown.Height = txtNm1.Height


'frmField.BackColor = vbBlue
End Sub
Private Sub imgButton2_Click()
'txtDown.Top = cmdCredit.Top
'txtDown.Left = cmdCredit.Left
txtDown.SetFocus

Hit = 0
Race = 1
Lace = 0
imgAccelR.Visible = False
imgAccelL.Visible = False
If Restart = 1 Then
Strok = 1     'for incr. speed after hit
Tim = 0    ' control tim.interval
Vin1 = 0   'count no of rounds win
Vin2 = 0
Final = 0
Sp1 = 0
Sp2 = 0
shpBall.Visible = True
imgP2.Visible = True
imgP1.Visible = True
 txtNm1.Enabled = False
 txtNm2.Enabled = False
imgP2.Height = 1200
'imgP2.Width = 170
imgP1.Height = 1200
'imgP1.Width = 170
txtVinner.Visible = False
'txtDown.Visible = True
'hide txtDown behind credit
txtDown.Top = picScroll.Top
txtDown.Left = picScroll.Left
txtDown.Width = picScroll.Width
imgNull1.Visible = True
imgNull2.Visible = True
imgOne1.Visible = False
imgOne2.Visible = False
imgTwo1.Visible = False
imgTwo2.Visible = False
imgOne.Visible = True
imgTwo.Visible = False
imgThree.Visible = False
timBall.Enabled = True
timBall.Interval = 150
imgNull1.Left = 3840
imgNull1.Top = 120
imgNull2.Left = 7920
imgNull2.Top = 120
frmField.BackColor = &HC0FFC0
imgVin.Visible = False
imgResult.Visible = False
Restart = 0
End If


imgButton2.Visible = False
txtDown.Top = picScroll.Top
txtDown.Left = picScroll.Left
If Final = 1 Then
' txtStart.Text = " PLAYER 1 WINS"
 txtDown.Top = imgResult.Top + imgResult.Height
txtDown.Left = imgResult.Left

 txtDown.Text = "START ANOTHER GAME"
 imgButton1.Visible = True
 Final = 0
 Vin1 = 0
 Vin2 = 0
 Restart = 1
ElseIf Final = 2 Then
 'txtStart.Text = " PLAYER 2 WINS"
 'txtStart.Visible = True
 'txtStart.Visible = True
 txtDown.Text = "START ANOTHER GAME"
 imgButton1.Visible = True
 Final = 0
 Vin1 = 0
 Vin2 = 0
 Restart = 1
 'End If
ElseIf Vin1 = 1 Then
  imgOne1.Visible = True
  imgNull1.Visible = False
  imgOne1.Left = 3840
  imgOne1.Top = 120
  If Vin2 = 0 Then
    imgOne2.Visible = False
    imgTwo2.Visible = False
    imgNull2.Left = 7920
    imgNull2.Top = 120
    imgOne.Visible = False
    imgTwo.Visible = True

  ElseIf Vin2 = 1 Then
    imgNull2.Visible = False
    imgOne2.Visible = True
    imgTwo2.Visible = False
    imgOne2.Left = 7920
    imgOne2.Top = 120
    imgTwo.Visible = False
    imgThree.Visible = True
 
   End If
  Sp1 = 0
  Sp2 = 0
  Strok = 0
  timBall.Interval = 100
  frmField.BackColor = &HFFC0C0
    txtDown.Top = picScroll.Top
  shpBall.Visible = True
   txtVinner.Visible = False
 txtVinner.Text = ""
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
   imgNull2.Visible = True
   imgOne2.Visible = False
   imgTwo2.Visible = False
   imgNull2.Left = 7920
   imgNull2.Top = 120
   'txtDown.Text = "PLAYER 1 WINS"
   'txtStart.Visible = True
  ElseIf Vin2 = 1 Then
    imgNull2.Visible = False
    imgOne2.Visible = True
    imgTwo2.Visible = False
     imgOne2.Left = 7920
    imgOne2.Top = 120
   'txtDown.Text = " PLAYER 1 WINS"
    'txtStart.Visible = True
   End If
  Sp1 = 0
  Sp2 = 0
  Strok = 0
  timBall.Interval = 70
  frmField.BackColor = &HFF8080
   txtVinner.Visible = False
 txtVinner.Text = ""
  imgTwo.Visible = False
  imgThree.Visible = True
  txtDown.Top = picScroll.Top
  shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False


ElseIf Vin2 = 1 Then
imgOne2.Visible = True
imgNull2.Visible = False
imgOne2.Left = 7920
imgOne2.Top = 120
If Vin1 = 0 Then
 imgNull1.Visible = True
 imgOne1.Visible = False
 imgTwo1.Visible = False
 imgNull1.Left = 3840
 imgNull1.Top = 120
 imgOne.Visible = False
 imgTwo.Visible = True
ElseIf Vin1 = 1 Then
 imgNull1.Visible = False
 imgOne1.Visible = True
 imgTwo1.Visible = False
 imgOne1.Left = 3840
 imgOne1.Top = 120
 imgTwo.Visible = False
 imgThree.Visible = True
 End If
Sp1 = 0
Sp2 = 0
Strok = 0
timBall.Interval = 100
frmField.BackColor = &HFFC0C0
 txtVinner.Visible = False
 txtVinner.Text = ""
txtDown.Top = picScroll.Top
shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False
ElseIf Vin2 = 2 Then
  imgTwo2.Visible = True
  imgOne2.Visible = False
  imgTwo2.Left = 7920
  imgTwo2.Top = 120
  If Vin2 = 0 Then
    imgNull2.Visible = True
    imgOne2.Visible = False
    imgTwo2.Visible = False
    imgNull1.Left = 3840
    imgNull1.Top = 120
    imgOne.Visible = False
    imgTwo.Visible = True
    'txtDown.Text = " PLAYER 2 WINS"
    'txtStart.Visible = True
   ElseIf Vin2 = 1 Then
    imgNull2.Visible = False
    imgOne2.Visible = True
    imgTwo2.Visible = False
    imgOne1.Left = 3840
    imgOne1.Top = 120
    imgTwo.Visible = False
    imgThree.Visible = True
    'txtDown.Text = " PLAYER 2 WINS"
    'txtStart.Visible = True
   End If
 Sp1 = 0
 Sp2 = 0
 Strok = 0
 timBall.Interval = 70
 frmField.BackColor = &HC0C0C08
 'imgTwo.Visible = False
 'imgThree.Visible = True
  txtVinner.Visible = False
 txtVinner.Text = ""
 txtDown.Top = picScroll.Top
 shpBall.Visible = True
  imgP1.Visible = True
  imgP2.Visible = True
  imgResult.Visible = False
  imgVin.Visible = False
Else
imgButton2.Visible = False
Nm1 = txtNm1.Text
 Nm2 = txtNm2.Text
imgStart.Visible = False
txtNm1.Enabled = False
txtNm1.Alignment = 2
txtNm1.Top = imgBot.Top
txtNm1.Left = imgBot.Left
txtNm2.Enabled = False
txtNm2.Top = imgBot.Top
txtNm2.Alignment = 2
txtNm2.Left = imgBot.Width - txtNm2.Width
timBall.Enabled = True
shpBall.Visible = True
imgP2.Visible = True
imgP1.Visible = True
txtDown.Top = picScroll.Top
txtDown.Left = picScroll.Left
frmField.BackColor = &HC0FFC0

'txtDown.Width = txtNm1.Width
'txtDown.Height = 285 'txtNm1.Height
End If

End Sub

Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)


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



Private Sub txtDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton1.Visible = True
imgButton2.Visible = False
End Sub

Private Sub txtStart_Click()
'End
Strok = 1     'for incr. speed after hit
Tim = 0    ' control tim.interval
Vin1 = 0   'count no of rounds win
Vin2 = 0
Final = 0
Sp1 = 0
Sp2 = 0
shpBall.Visible = True
imgP2.Visible = True
imgP1.Visible = True
 
imgP2.Height = 1200
imgP2.Width = 170
imgP1.Height = 1200
imgP1.Width = 170
txtStart.Visible = False
txtDown.Visible = True
'hide txtDown behind credit
txtDown.Top = picScroll.Top
txtDown.Left = picScroll.Left
txtDown.Width = picScroll.Width
imgNull1.Visible = True
imgNull2.Visible = True
imgOne1.Visible = False
imgOne2.Visible = False
imgTwo1.Visible = False
imgTwo2.Visible = False
imgOne.Visible = True
imgTwo.Visible = False
imgThree.Visible = False
timBall.Enabled = True
timBall.Interval = 150
imgNull1.Left = 3840
imgNull1.Top = 120
imgNull2.Left = 7920
imgNull2.Top = 120
frmField.BackColor = &HC0FFC0
imgVin.Visible = False
imgResult.Visible = False


End Sub
