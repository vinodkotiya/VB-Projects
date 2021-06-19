VERSION 5.00
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALVIN"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   Icon            =   "FRMCAL~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdXpy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "x^y"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdMr 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MR"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdMp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "M+"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdCe 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CE"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdReci 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1/x"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox picF6 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "FRMCAL~1.frx":030A
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "--->"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdSqrt 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Sqrt."
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdPm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "UniversalMath1 BT"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4200
      Width           =   495
   End
   Begin VB.Timer timFace 
      Interval        =   200
      Left            =   120
      Top             =   480
   End
   Begin VB.PictureBox picF5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "FRMCAL~1.frx":0614
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picF4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "FRMCAL~1.frx":091E
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picF3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "FRMCAL~1.frx":0C28
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picF2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "FRMCAL~1.frx":0F32
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picF1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "FRMCAL~1.frx":123C
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPer 
      BackColor       =   &H00C0FFC0&
      Caption         =   "%"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdNul 
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdFor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdFiv 
      BackColor       =   &H00FFC0C0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdSix 
      BackColor       =   &H00FFC0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdMul 
      BackColor       =   &H00C0FFC0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "UniversalMath1 BT"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdMin 
      BackColor       =   &H00C0FFC0&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdPlu 
      BackColor       =   &H00C0FFC0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdNin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdAt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdSev 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdThr 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdTwo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdOne 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdEqu 
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdDec 
      BackColor       =   &H00FFC0C0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "UniversalMath1 BT"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdDiv 
      BackColor       =   &H00C0FFC0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdCredit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Credit"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FF00FF&
      Caption         =   "On"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtLcd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FRMCAL~1.frx":1546
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   0
      Picture         =   "FRMCAL~1.frx":1582
      Top             =   -120
      Width           =   4500
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data1 As Single
Dim Data2 As Single
Dim Res As Single
Dim Op As Integer
Dim Flag As Integer       'used for taking Data2 or identify Data1/2
Dim Ref As Integer       'refreshing LCD after =
Dim F As Integer
Dim Lcd As Integer        'display text when off
Dim Pm As Single          'For storing data when +/- click
Dim ResMem As Single      'For M+

Private Sub cmdAt_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "8"
End Sub

Private Sub cmdBack_Click()
Beep
If cmdSwitch.Caption = "Off" Then
txtLcd.Text = txtLcd.Text \ 10
 If txtLcd.Text = "0" Then
 txtLcd.Text = ""
 End If
End If
End Sub

Private Sub cmdCe_Click()
Beep
txtLcd.Text = ""
End Sub

Private Sub cmdCredit_Click()
Beep
frmCredit.Visible = True
End Sub

Private Sub cmdDec_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "."
End Sub

Private Sub cmdDiv_Click()
Beep
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 4
End If
End Sub

Private Sub cmdEsc_Click()
End
End Sub

Private Sub cmdMp_Click()
ResMem = ResMem + Val(txtLcd.Text)
End Sub

Private Sub cmdMr_Click()
txtLcd.Text = Str(ResMem)
End Sub

Private Sub cmdPer_Click()
Beep
If Flag = 1 Then
Data2 = Val(txtLcd.Text)
Flag = 0
End If
Res = (Data1 * Data2) / 100
txtLcd.Text = Str(Res)
Ref = 1
End Sub

Private Sub cmdPlu_Click()
Beep
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 1
End If
End Sub


Private Sub cmdEqu_Click()
Beep
If Flag = 1 Then
Data2 = Val(txtLcd.Text)
Flag = 0     'next will Data1
End If
If Op = 1 Then
Res = Data1 + Data2
ElseIf Op = 2 Then
Res = Data1 - Data2
ElseIf Op = 3 Then
Res = Data1 * Data2
ElseIf Op = 4 Then
  If Data2 > 0 Then
  Res = Data1 / Data2
  ElseIf Data2 < 0 Then
  Res = Data1 / Data2
  Else
  txtLcd.Text = " CAN NOT DIVIDE BY ZERO U FOOL"
  End If
ElseIf Op = 5 Then
Res = Data1 ^ Data2
End If
If txtLcd.Text = " CAN NOT DIVIDE BY ZERO U FOOL" Then
Ref = 1
Else
txtLcd.Text = Str(Res)
Ref = 1
End If
End Sub

Private Sub cmdFiv_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "5"
End Sub

Private Sub cmdFor_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "4"
End Sub

Private Sub cmdMin_Click()
Beep
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 2
End If
End Sub

Private Sub cmdMul_Click()
Beep
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 3
End If

End Sub

Private Sub cmdNin_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "9"
End Sub

Private Sub cmdNul_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "0"
End Sub

Private Sub cmdPm_Click()
Beep
Pm = Val(txtLcd.Text) * -1
txtLcd.Text = Str(Pm)
End Sub

Private Sub cmdReci_Click()
Beep
If cmdSwitch.Caption = "Off" Then
Res = 1 / txtLcd.Text
txtLcd.Text = Str(Res)
Ref = 1
End If
End Sub

Private Sub cmdSqrt_Click()
Beep
Res = Sqr(Val(txtLcd.Text))
txtLcd.Text = Str(Res)
Ref = 1
End Sub

Private Sub cmdSwitch_Click()
Beep
If cmdSwitch.Caption = "Off" Then
cmdSwitch.Caption = "On"
cmdSwitch.BackColor = &HFF00FF
Lcd = 1
F = 0
Else
F = 7
txtLcd.Text = ""
Lcd = 0
cmdSwitch.Caption = "Off"
cmdSwitch.BackColor = &H80FF80
End If
End Sub


Private Sub cmdOne_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "1"
End Sub


Private Sub cmdSev_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "7"
End Sub

Private Sub cmdSix_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "6"
End Sub

Private Sub cmdThr_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "3"
End Sub

Private Sub cmdTwo_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "2"
End Sub







Private Sub cmdXpy_Click()
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 5
End If
End Sub

Private Sub Form_Load()
Flag = 0
Ref = 0   'For Refreshing Lcd
Lcd = 1
ResMem = 0  'For M+
picF6.Visible = True
End Sub













Private Sub timFace_Timer()
If Lcd = 0 Then
'txtLcd.Text = ""
ElseIf Lcd < 15 Then
txtLcd.Text = "  HELLO!    I AM CALVIN.................."
Lcd = Lcd + 1
ElseIf Lcd = 105 Then
Lcd = 1
ElseIf Lcd > 90 Then
txtLcd.Text = " COPY & DISTRIBUTE FREE4ALL"
Lcd = Lcd + 1
ElseIf Lcd > 75 Then
txtLcd.Text = "  FREEWARE / FUNWARE........"
Lcd = Lcd + 1
ElseIf Lcd > 60 Then
txtLcd.Text = " ALL RIGHTS UNRESERVED ......."
Lcd = Lcd + 1
ElseIf Lcd > 45 Then
txtLcd.Text = "(c) COPYWRITE 2002 VINOD KOTIYA"
Lcd = Lcd + 1
ElseIf Lcd > 30 Then
txtLcd.Text = " CLICK 'On' To START THE CALVIN"
Lcd = Lcd + 1
ElseIf Lcd > 14 Then
txtLcd.Text = " YOUR MINI CALCULATOR............"
Lcd = Lcd + 1
End If
If F = 1 Then
picF2.Visible = False
picF1.Visible = True
F = 2
ElseIf F = 2 Then
picF1.Visible = False
picF2.Visible = True
F = 3
ElseIf F = 3 Then
picF2.Visible = False
picF3.Visible = True
F = 4
ElseIf F = 4 Then
picF3.Visible = False
picF4.Visible = True
F = 5
ElseIf F = 5 Then
picF4.Visible = False
picF5.Visible = True
F = 6
ElseIf F = 6 Then
picF5.Visible = False
picF6.Visible = True
F = 7
ElseIf F = 7 Then
picF6.Visible = False
picF5.Visible = True
F = 8
ElseIf F = 8 Then
picF5.Visible = False
picF4.Visible = True
F = 9
ElseIf F = 9 Then
picF4.Visible = False
picF3.Visible = True
F = 10
ElseIf F = 10 Then
picF3.Visible = False
picF2.Visible = True
F = 1
End If
End Sub

Private Sub txtLcd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 107 Then
  If Flag = 0 Then
  Data1 = Val(txtLcd.Text)
  Flag = 1
  txtLcd.Text = ""
  Op = 1
  End If
  
ElseIf KeyCode = 187 Then
  If Flag = 1 Then
  Data2 = Val(txtLcd.Text)
  Flag = 0
  End If
  If Op = 1 Then
  Res = Data1 + Data2
  ElseIf Op = 2 Then
  Res = Data1 - Data2
  ElseIf Op = 3 Then
  Res = Data1 * Data2
  ElseIf Op = 4 Then
  If Data2 = 0 Then
  txtLcd.Text = "error"
  Else
  Res = Data1 / Data2
  End If
'ElseIf Op = 5 Then
'Res = (Data1 * Data2) / 100
End If
txtLcd.Text = Str(Res)
Ref = 1

End If
End Sub

