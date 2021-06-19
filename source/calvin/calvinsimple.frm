VERSION 5.00
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALVIN"
   ClientHeight    =   6135
   ClientLeft      =   4380
   ClientTop       =   2430
   ClientWidth     =   4395
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   Icon            =   "calvinsimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4395
   Begin VB.TextBox txtDate 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   80
      Width           =   1095
   End
   Begin VB.CommandButton cmdCredit 
      BackColor       =   &H00FFFF80&
      Caption         =   "CREDIT"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto Copy Result"
      Height          =   495
      Left            =   1320
      MaskColor       =   &H008080FF&
      TabIndex        =   35
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Set on Top of All Windows"
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlu 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Picture         =   "calvinsimple.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdCon 
      BackColor       =   &H0000FFFF&
      Caption         =   "Scientific"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdXpy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "x^y"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdMr 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MR"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdMp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "M+"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdCe 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CE"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdReci 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1/x"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3720
      Width           =   615
   End
   Begin VB.PictureBox picF6 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "calvinsimple.frx":2478
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "--->"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdSqrt 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Sqrt."
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdPm 
      BackColor       =   &H00C0FFC0&
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
      Picture         =   "calvinsimple.frx":2782
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4440
      Width           =   615
   End
   Begin VB.Timer timFace 
      Interval        =   200
      Left            =   3000
      Top             =   720
   End
   Begin VB.PictureBox picF5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "calvinsimple.frx":2846
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picF4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "calvinsimple.frx":2B50
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picF3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "calvinsimple.frx":2E5A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picF2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "calvinsimple.frx":3164
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picF1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "calvinsimple.frx":346E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPer 
      BackColor       =   &H00C0FFC0&
      Caption         =   "%"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2280
      Width           =   615
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdMul 
      BackColor       =   &H00FF0000&
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
      Left            =   3720
      Picture         =   "calvinsimple.frx":3778
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdMin 
      BackColor       =   &H00FF0000&
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
      Left            =   3720
      Picture         =   "calvinsimple.frx":3F26
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdDec 
      BackColor       =   &H00FFC0C0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdDiv 
      BackColor       =   &H00FF0000&
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
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "calvinsimple.frx":46D4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FF00FF&
      Caption         =   "On"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtLcd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "calvinsimple.frx":4E82
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image imgScroll 
      Height          =   450
      Left            =   480
      Picture         =   "calvinsimple.frx":4EBE
      Top             =   5640
      Width           =   12000
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   0
      Picture         =   "calvinsimple.frx":16840
      Top             =   -120
      Width           =   4500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   555
      Left            =   0
      Top             =   5580
      Width           =   4455
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
Dim Autocopycode, dashamlav As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40



Private Sub Check1_Click()
If Check1.Value = vbChecked Then
Dim retValue, retsci As Long
    'Load Form1
    retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 260, 0, _
               285, 432, SWP_SHOWWINDOW)
    If cmdCon.Caption = "Simple" Then
     retsci = SetWindowPos(frmSci.hwnd, HWND_TOPMOST, 540, 100, _
               185, 285, SWP_SHOWWINDOW)
     End If
ElseIf Check1.Value = vbUnchecked Then
   Dim reetValue As Long
   
    reetValue = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 200, 10, _
               285, 432, SWP_SHOWWINDOW)
    If cmdCon.Caption = "Simple" Then
     retsci = SetWindowPos(frmSci.hwnd, HWND_NOTOPMOST, 480, 100, _
               185, 285, SWP_SHOWWINDOW)
     End If
End If

End Sub


Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSci.Left = frmCal.Left + frmCal.Width
frmSci.Top = frmCal.Top + 1200

End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
  Autocopycode = True
 Else
 Autocopycode = False
End If

End Sub

Private Sub cmdAt_Click()
Beep
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "8"
End Sub

Private Sub cmdBack_Click()
Dim l As Integer
Beep
If cmdSwitch.Caption = "Off" Then
l = Len(txtLcd.Text)
If l > 0 Then
 txtLcd.Text = Mid$(txtLcd.Text, 1, l - 1)
  If txtLcd.Text = "0" Then
  txtLcd.Text = ""
  End If
 Else
 MsgBox "There is nothing to delete"
 End If
End If
End Sub

Private Sub cmdCe_Click()
Beep
txtLcd.Text = ""
End Sub

Private Sub cmdCon_Click()
If cmdCon.Caption = "Scientific" Then
Load frmSci
frmSci.Visible = True
cmdCon.Caption = "Simple"
Else

frmSci.Visible = False
Unload frmSci
'frmSci.Enabled = False
cmdCon.Caption = "Scientific"
'frmSci.ActiveControl = False
End If
End Sub

Private Sub cmdCredit_Click()
Beep
frmCredit.Visible = True
End Sub

Private Sub cmdDec_Click()
Dim ret As Boolean
Beep
If cmdSwitch.Caption = "Off" Then
 If txtLcd.Text = "" Then
  txtLcd.Text = "0"
 ElseIf txtLcd.Text = "." Then
  txtLcd.Text = "0"
 ElseIf txtLcd.Text = "0." Then
  txtLcd.Text = "0"
 End If

ret = Checkpoint()
If ret = False Then
 txtLcd.Text = txtLcd.Text & "."
 Else
 MsgBox "Don't you know that only one decimal should be placed in a number"
End If
End If
End Sub
Function Checkpoint() As Boolean
Dim l, i As Integer
Dim sesh, num As String
num = txtLcd.Text
l = Len(txtLcd.Text)
For i = 1 To l
 sesh = Mid$(num, i, 1)
 If sesh = "." Then
  Checkpoint = True  'ek dashamlav mila
  Exit Function
 End If
Next i
Checkpoint = False
End Function
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
If cmdSwitch.Caption = "On" Then
 If timFace.Interval < 11 Then
  timFace.Interval = 20
 Else
  timFace.Interval = timFace.Interval - 10
 End If
Else

 If Flag = 0 Then
  Data1 = Val(txtLcd.Text)
  Flag = 1   'data1 geted
  txtLcd.Text = ""
  Op = 1
 End If
End If

End Sub


Private Sub cmdEqu_Click()
On Error GoTo vinerror
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
ElseIf (Res > (4.94065645841247E-324) And Res < (1.7E+308)) Or (Res > (-4.94065645841247E-324) And Res < (-1.7E+308)) Then
txtLcd.Text = Str(Res)
Ref = 1
End If
dashamlav = False
 Exit Sub
vinerror:
 MsgBox "An Overflow occured "
 txtLcd.Text = " "
 Exit Sub
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
Private Sub cmdPlus_Click()
Beep
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 1
End If

End Sub

Private Sub cmdMin_Click()
Beep
If cmdSwitch.Caption = "On" Then
 If timFace.Interval > 400 Then
  timFace.Interval = 390
 Else
 timFace.Interval = timFace.Interval + 10
 End If
Else

 If Flag = 0 Then
  Data1 = Val(txtLcd.Text)
  Flag = 1
  txtLcd.Text = ""
  Op = 2
  End If
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
 If Val(txtLcd.Text) = 0 Then
  MsgBox "Can't Divide By zero"
 Else
  Res = 1 / txtLcd.Text
  txtLcd.Text = Str(Res)
  Ref = 1
 End If
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
txtLcd.MaxLength = 25
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
dashamlav = False
txtDate.Text = Now 'Date & Time
End Sub

















Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSci.Left = frmCal.Left + frmCal.Width
frmSci.Top = frmCal.Top + 1200

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSci
Unload frmCredit
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSci.Left = frmCal.Left + frmCal.Width
frmSci.Top = frmCal.Top + 1200
If Autocopycode = True Then
 Clipboard.Clear
 Clipboard.SetText txtLcd.Text, vbCFText
End If

End Sub

Private Sub timFace_Timer()
txtDate.Text = Now 'txtDate.Text = Date & Time()
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
''for refreshing screen after frmsci done
If frmSci.txtRefresh.Text = "1" Then
Ref = 1
frmSci.txtRefresh.Text = ""
End If
''scroll
If (imgScroll.Left + imgScroll.Width - 1000) > frmCal.Left Then
 imgScroll.Left = imgScroll.Left - 200
ElseIf (imgScroll.Left + imgScroll.Width - 1000) < frmCal.Left Then
 imgScroll.Left = (frmCal.Left + frmCal.ScaleWidth)
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

