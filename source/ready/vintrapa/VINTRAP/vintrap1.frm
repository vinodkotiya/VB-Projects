VERSION 5.00
Begin VB.Form frmField 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VINTRAP (DEMO VER.)"
   ClientHeight    =   8190
   ClientLeft      =   2085
   ClientTop       =   930
   ClientWidth     =   11880
   Icon            =   "vintrap1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "vintrap1.frx":0ECA
   Picture         =   "vintrap1.frx":101C
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdMore 
      Caption         =   "More.."
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame fraScore 
      Caption         =   "Select Score Limit"
      Height          =   975
      Left            =   4200
      TabIndex        =   19
      Top             =   5520
      Width           =   4455
      Begin VB.OptionButton score4 
         Caption         =   "60"
         Height          =   495
         Left            =   3120
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton score3 
         Caption         =   "75"
         Height          =   495
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton score2 
         Caption         =   "50"
         Height          =   495
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton score1 
         Caption         =   "25"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtVel 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "speed"
      ToolTipText     =   "show speed of ball  in centimeter per second"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1920
      MousePointer    =   1  'Arrow
      ScaleHeight     =   585
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   7680
      Width           =   8895
   End
   Begin VB.Frame Processor 
      BackColor       =   &H00CEFDFC&
      Caption         =   "Select Your MicroProcessor"
      Height          =   1095
      Left            =   4680
      TabIndex        =   11
      Top             =   1560
      Width           =   3255
      Begin VB.OptionButton P2 
         BackColor       =   &H00CEFDFC&
         Caption         =   "P-II"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "If your processsor is less then 600 MHz"
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton P3 
         BackColor       =   &H00CEFDFC&
         Caption         =   "P-III"
         Height          =   495
         Left            =   1080
         TabIndex        =   13
         ToolTipText     =   "If your processsor is between 600 MHz & 1.3 GHz"
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton P4 
         BackColor       =   &H00CEFDFC&
         Caption         =   "P-IV"
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         ToolTipText     =   "If your processsor is more then 1.3 GHz"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "v/s Player"
      Height          =   495
      Left            =   6720
      MouseIcon       =   "vintrap1.frx":1EE6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Click to Play against your friend"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "v/s Computer"
      Height          =   495
      Left            =   4680
      MouseIcon       =   "vintrap1.frx":21F0
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Click to Play against computer."
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Level 
      BackColor       =   &H00CEFDFC&
      Caption         =   "PLEASE SELECT A LEVEL"
      Height          =   975
      Left            =   4680
      MouseIcon       =   "vintrap1.frx":24FA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2640
      Width           =   3255
      Begin VB.OptionButton Option1 
         BackColor       =   &H00CEFDFC&
         Caption         =   "Fast"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Too Easy mode against the computer"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00CEFDFC&
         Caption         =   "Faster"
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         ToolTipText     =   "Normal mode against the computer"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00CEFDFC&
         Caption         =   "Fastest"
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         ToolTipText     =   "Hardest mode against the computer"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer timFun 
      Interval        =   50
      Left            =   1320
      Top             =   3360
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
      Left            =   3840
      TabIndex        =   5
      Text            =   "VINOD"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtNm2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4320
      MouseIcon       =   "vintrap1.frx":2804
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Text            =   "Player2"
      ToolTipText     =   "Name of Player2"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtNm1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4300
      MouseIcon       =   "vintrap1.frx":2B0E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Text            =   "Player1"
      ToolTipText     =   "Name of Player1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtP2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Score of Player2"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtP1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Score of Player1"
      Top             =   615
      Width           =   615
   End
   Begin VB.Timer timBall 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   600
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
      Left            =   2760
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Text            =   "CLICK ON STAR TO BEGIN THE GAME"
      Top             =   3960
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   855
      Left            =   360
      OleObjectBlob   =   "vintrap1.frx":2E18
      SourceDoc       =   "F:\VB PROJECTS\ready\vintrap\readme.txt"
      TabIndex        =   18
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image imgReset 
      Height          =   330
      Left            =   11040
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap1.frx":5030
      Top             =   8040
      Width           =   750
   End
   Begin VB.Image imgSpeed 
      Height          =   675
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap1.frx":5604
      Top             =   7680
      Width           =   1950
   End
   Begin VB.Image shpBall2 
      Height          =   480
      Left            =   1320
      Picture         =   "vintrap1.frx":9B2E
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMe 
      Height          =   720
      Left            =   8520
      Picture         =   "vintrap1.frx":9E38
      Top             =   6720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image shpBall 
      Height          =   480
      Left            =   840
      Picture         =   "vintrap1.frx":BB02
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image cmdCredit 
      Height          =   720
      Left            =   8160
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap1.frx":BE0C
      ToolTipText     =   "Author:-Vinod Kotiya"
      Top             =   6720
      Width           =   720
   End
   Begin VB.Image imgButton2 
      Height          =   720
      Left            =   4200
      MouseIcon       =   "vintrap1.frx":CCD6
      MousePointer    =   99  'Custom
      Picture         =   "vintrap1.frx":CFE0
      ToolTipText     =   "Click every time on this green star to start/continue the game"
      Top             =   3240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgButton1 
      Height          =   720
      Left            =   4200
      Picture         =   "vintrap1.frx":ECAA
      Top             =   3240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgAccelL 
      Height          =   480
      Left            =   5880
      Picture         =   "vintrap1.frx":10974
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAccelR 
      Height          =   480
      Left            =   5400
      Picture         =   "vintrap1.frx":11362
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   1785
      Left            =   3120
      MouseIcon       =   "vintrap1.frx":11D50
      MousePointer    =   99  'Custom
      Picture         =   "vintrap1.frx":1205A
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgTwo2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "vintrap1.frx":1C2A8
      ToolTipText     =   "Round Results"
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgTwo1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "vintrap1.frx":1D17A
      ToolTipText     =   "Round Results"
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgThree 
      Height          =   330
      Left            =   6960
      Picture         =   "vintrap1.frx":1E04C
      ToolTipText     =   "You are currently playing this round"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgResult 
      Height          =   750
      Left            =   3840
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap1.frx":1E6F6
      Top             =   3000
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgVin 
      Height          =   750
      Left            =   480
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap1.frx":229B8
      Top             =   2160
      Visible         =   0   'False
      Width           =   10800
   End
   Begin VB.Image imgTwo 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   6960
      Picture         =   "vintrap1.frx":2BA9A
      ToolTipText     =   "You are currently playing this round"
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgOne2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "vintrap1.frx":2C144
      ToolTipText     =   "Round Results"
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgOne1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "vintrap1.frx":2D016
      ToolTipText     =   "Round Results"
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgNull2 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   7920
      Picture         =   "vintrap1.frx":2DEE8
      ToolTipText     =   "Round Results"
      Top             =   120
      Width           =   840
   End
   Begin VB.Image imgNull1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "vintrap1.frx":2EDBA
      ToolTipText     =   "Round Results"
      Top             =   120
      Width           =   840
   End
   Begin VB.Image imgBot 
      Height          =   1200
      Left            =   0
      MousePointer    =   3  'I-Beam
      Picture         =   "vintrap1.frx":2FC8C
      Top             =   6480
      Width           =   12000
   End
   Begin VB.Image imgOne 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   6960
      Picture         =   "vintrap1.frx":3FACE
      ToolTipText     =   "You are currently playing this round"
      Top             =   360
      Width           =   435
   End
   Begin VB.Image imgCap 
      Height          =   1050
      Left            =   -10
      MousePointer    =   1  'Arrow
      Picture         =   "vintrap1.frx":40178
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image imgP1 
      Height          =   1920
      Left            =   1440
      Picture         =   "vintrap1.frx":4E07A
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgP2 
      Height          =   1920
      Left            =   10080
      Picture         =   "vintrap1.frx":4ECBC
      Top             =   4440
      Width           =   240
   End
   Begin VB.Shape shpSq 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   735
      Left            =   10800
      Top             =   7560
      Width           =   1215
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
Dim Leveler As Integer  ' For hardness
Dim Ballchange As Integer
Dim Procesor As Integer
Dim d1 As Integer, d2 As Integer, Vel As Integer, Velocity As Currency 'determine velocity
Dim Score As Integer


Private Sub cmdCredit_Click()


'frmField.Visible = False
timFun.Interval = timBall.Interval
timBall.Interval = 0
'restorespeed = 1
'Tim = 1
Load frmCredit
frmCredit.Visible = True
End Sub



Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub File1_Click()

End Sub

Private Sub cmdCredit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMe.Visible = True
End Sub

Private Sub cmdMore_Click()
Load frmAsk
frmAsk.ScaleHeight = 2355
frmAsk.ScaleWidth = 4690
frmAsk.lblId.Visible = True
frmAsk.txtId.Visible = True
frmAsk.cmdBack.Visible = True
frmAsk.cmdNext.Visible = True
frmAsk.cmdCredit.Visible = True
frmAsk.lblAsk.Visible = False
frmAsk.imgYes.Visible = False
frmAsk.imgNo.Visible = False
frmAsk.Visible = True

End Sub

Private Sub Command1_Click()
 Com = 1
 imgStart.Visible = True
 txtNm1.Visible = True
 txtNm1.Text = "VINOD KOTIYA"
 txtNm2.Visible = True
 imgButton1.Visible = True
 txtDown.Visible = True
 Level.Visible = False
 Option1.Visible = False
 Option2.Visible = False
 Option3.Visible = False
 Processor.Visible = False
 P2.Visible = False
 P3.Visible = False
 P4.Visible = False
 Command1.Visible = False
 Command2.Visible = False
 fraScore.Visible = False
score1.Visible = False
score2.Visible = False
score3.Visible = False
score4.Visible = False

End Sub

Private Sub Command2_Click()
 Com = 0
 imgStart.Visible = True
 txtNm1.Visible = True
 txtNm2.Visible = True
 imgButton1.Visible = True
 txtDown.Visible = True
  Level.Visible = False
 Option1.Visible = False
 Option2.Visible = False
 Option3.Visible = False
 Processor.Visible = False
 P2.Visible = False
 P3.Visible = False
 P4.Visible = False
 Command1.Visible = False
 Command2.Visible = False
 fraScore.Visible = False
score1.Visible = False
score2.Visible = False
score3.Visible = False
score4.Visible = False

End Sub


Private Sub Command3_Click()
End Sub

Private Sub imgReset_Click()

Load frmAsk
frmAsk.ScaleHeight = 1290
frmAsk.ScaleWidth = 4690
timFun.Interval = timBall.Interval
timBall.Interval = 0
frmAsk.lblId.Visible = False
frmAsk.txtId.Visible = False
frmAsk.cmdBack.Visible = False
frmAsk.cmdNext.Visible = False
frmAsk.cmdCredit.Visible = False
frmAsk.Visible = True
'Unload Me
End Sub

Private Sub Form_Load()
 Option2.Value = True     ''default faster
 P3.Value = True         ''default P-3
 score1.Value = True
 frmField.ScaleHeight = 8190
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
' txtDown.Visible = True
 imgP2.Visible = False
 Cap = 1050    'used for imgCap.Height
 Bot = imgBot.Height + picScroll.Height - 100
 
 Strok = 1     'for incr. speed after hit
 Tim = 0    ' control tim.interval
 Vin1 = 0   'count no of rounds win
 Vin2 = 0
 Ballchange = 1
 'imgP2.Height = 1200
 'imgP2.Width = 170
 'imgP1.Height = 1200
 'imgP1.Width = 170
imgP1.Left = 1550
imgP2.Left = frmField.ScaleWidth - 1550 - imgP2.Width
imgBot.Top = (frmField.Height - imgBot.Height - picScroll.Height)
'picScroll.Width = frmField.Width
picScroll.Top = frmField.Height - picScroll.Height
imgSpeed.Top = imgBot.Top + imgBot.Height
shpSq.Top = imgSpeed.Top
imgReset.Top = imgSpeed.Top + 100
txtVel.Left = imgSpeed.Left + imgSpeed.Width
txtVel.Top = txtVel.Top + 70
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
 Temp = 0
 'Com = 1
'Procesor = 100  'P-4 150 & div 80 'P-3 120 div 70'P-2 100 div 60
'Div = 60     'for dividing ballx bally
Vel = 0
Score = 25
 End Sub

Private Sub imgRod_Click()
End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload frmAsk
Unload frmCredit
End Sub

Private Sub imgBot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMe.Visible = False
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



Private Sub Option1_Click()
Leveler = 20
End Sub

Private Sub Option2_Click()
Leveler = 40
End Sub

Private Sub Option3_Click()
Leveler = 60

End Sub


Private Sub P2_Click()
Div = 60
Procesor = 100
End Sub

Private Sub P3_Click()
Div = 70
Procesor = 120
End Sub

Private Sub P4_Click()
Div = 80
Procesor = 150
End Sub

Private Sub score1_Click()
Score = 25
End Sub

Private Sub score2_Click()
Score = 50
End Sub

Private Sub score3_Click()
Score = 75
End Sub

Private Sub score4_Click()
Score = 100
End Sub

Private Sub timBall_Timer()
 'txtDown.SetFocus
'Temp = txtDown_Click()

'Text1.Text = Str(timBall.Interval)
''determine the velocity
   Vel = Vel + 1  'initialy 0
   If Vel = 1 Then
   d1 = BallX
   End If
   If Vel = 2 Then
    d2 = BallX
    Velocity = ((d2 - d1) / 567)
    If timBall.Interval > 0 Then
        Velocity = (Velocity / timBall.Interval) * 1000
    End If
    If Velocity < 0 Then
    Velocity = Velocity * (-1)
    End If
   txtVel.Text = Str(Velocity) & " cm/sec"
   Vel = 0
  End If

''ani ball
Ballchange = Ballchange + 1
If Ballchange > 7 Then
 Ballchange = 1
End If
If Ballchange > 5 Then
 shpBall2.Top = BallY
 shpBall2.Left = BallX
 shpBall2.Visible = True
 shpBall.Visible = False
 'Ballchange = Ballchange + 1
ElseIf Ballchange > 0 Then
 shpBall.Visible = True
 shpBall2.Visible = False
End If

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
 If timBall.Interval > 10 Then ''20
  If Strok Mod 2 = 0 Then     ' incr speed when strike
  timBall.Interval = timBall.Interval - 10
  End If
  Else
   timBall.Interval = 10
 End If
 
Tim = 0     'becomes 1 when strike
End If      'tim ended
If incrspeed = 1 Then
 Temp = timBall.Interval ' stores the current speed & make speed 10 when strike restore speed & next accelarator will display
 timBall.Interval = 3
 restorespeed = 1
 incrspeed = 0
' txtDown.Text = Str(Temp)
End If
shpBall.Top = BallY
'score
txtP1.Text = Str(Sp1)
txtP2.Text = Str(Sp2)

'change Round
If Sp1 = Score Then
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
 imgButton1.Left = imgResult.Left + 1680
 imgAccelR.Visible = False
 imgAccelL.Visible = False
 txtVel.Visible = False
 frmField.MousePointer = 0
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
   txtDown.Text = "     START SECOND ROUND"
   ElseIf Vin2 = 1 Then
    imgOne2.Visible = True
    imgNull2.Visible = False
    imgOne2.Top = imgVin.Top + imgVin.Height
   imgOne2.Left = imgResult.Left + imgResult.Width
   txtDown.Text = "       START FINAL ROUND"
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
  txtDown.Text = "CONGRATULATIONS " & Nm1 & " FOR WINNING"
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
ElseIf Sp2 = Score Then
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
  imgButton1.Left = imgResult.Left + 1680
 imgAccelR.Visible = False
 imgAccelL.Visible = False
  txtVel.Visible = False
 frmField.MousePointer = 0
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
     txtDown.Text = "         START SECOND ROUND"
    ElseIf Vin1 = 1 Then
     imgOne1.Visible = True
     imgNull1.Visible = False
     imgOne1.Top = imgOne2.Top
     imgOne1.Left = imgResult.Left - imgOne1.Width
     txtDown.Text = "         START FINAL ROUND"
    ElseIf Vin1 = 2 Then
     Final = 1
     End If
    'imgOne.Visible = False
    'imgTwo.Visible = True
  ElseIf Vin2 = 2 Then
    ' = txtDown.WidthimgNull1.Top + imgNull1.Height
     txtDown.Text = "CONGRATULATIONS " & Nm2 & " FOR WINNING"
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

    End If 'balldiry = -1 ended
     
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
    imgP1.Top = imgP1.Top - Leveler
   Else
    imgP1.Top = Cap
   End If
  ElseIf BallDiry = 1 Then
   If imgP1.Top < (frmField.ScaleHeight - imgP1.Height - Bot) Then
     imgP1.Top = imgP1.Top + Leveler
     Else
     imgP1.Top = (frmField.ScaleHeight - imgP1.Height - Bot)
    End If
  End If
End If     'End com
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMe.Visible = False
If Y < frmField.ScaleHeight - imgP2.Height Then
imgP2.Top = Y - 200
End If
If imgP2.Top > (frmField.ScaleHeight - Bot - imgP2.Height) Then
imgP2.Top = (frmField.ScaleHeight - Bot - imgP2.Height)
End If

End Sub

Private Sub txtDown_Click()
txtDown.Top = picScroll.Top
txtDown.Left = picScroll.Left
End Sub
Private Sub imgButton2_Click()
OLE1.Visible = False
cmdMore.Visible = False
frmField.MousePointer = 99
txtVel.Visible = True
txtDown.SetFocus
BallX = frmField.ScaleWidth / 2
BallY = frmField.ScaleHeight / 2
shpBall.Left = BallX
shpBall.Top = BallY
imgP1.Top = 3960
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
 ''txtNm1.Enabled = False
'' txtNm2.Enabled = False
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
timBall.Interval = Procesor ''''
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
  timBall.Interval = Procesor - 50  'incr speed in next round ''100
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
  timBall.Interval = Procesor - 90 'incr speed in next round
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
timBall.Interval = Procesor - 50 'incr speed in next round''100
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
 timBall.Interval = Procesor - 90 'incr speed in next round''70
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
txtNm1.BackColor = vbBlack
txtNm1.ForeColor = vbGreen
txtNm1.Alignment = 2
txtNm1.Top = imgBot.Top + 60
txtNm1.Left = imgBot.Left
txtNm2.BackColor = vbBlack
txtNm2.ForeColor = vbGreen
txtNm2.Top = imgBot.Top + 60
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
If Com = 0 Then  '' if opponent is not computer
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
End If 'end com

End Sub



Private Sub txtDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgButton1.Visible = True
imgButton2.Visible = False
End Sub


