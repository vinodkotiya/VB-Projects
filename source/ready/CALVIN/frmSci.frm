VERSION 5.00
Begin VB.Form frmSci 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCIENTIFIC"
   ClientHeight    =   3915
   ClientLeft      =   6435
   ClientTop       =   2820
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSci.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdRnd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rnd"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdRound 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Round"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdAbs 
      BackColor       =   &H00FFC0C0&
      Caption         =   "| Abs |"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdHex 
      BackColor       =   &H00FFFF00&
      Caption         =   "hex "
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      BackColor       =   &H0080FFFF&
      Caption         =   "cos"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdTan 
      BackColor       =   &H0080FFFF&
      Caption         =   "tan"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdLogn 
      BackColor       =   &H0080FF80&
      Caption         =   "ln"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdExp 
      BackColor       =   &H0080FF80&
      Caption         =   "exp"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H0080FF80&
      Caption         =   "log"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdFact 
      BackColor       =   &H008080FF&
      Caption         =   "n !"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdOct 
      BackColor       =   &H00FFFF00&
      Caption         =   "oct"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdSin 
      BackColor       =   &H0080FFFF&
      Caption         =   "sin"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtRefresh 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "frmSci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, r As Integer

Private Sub cmdAbs_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Abs(Val(frmCal.txtLcd.Text))
txtRefresh.Text = "1"
End If

End Sub

Private Sub cmdCombi_Click()

End Sub

Private Sub cmdCos_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Cos(Val(frmCal.txtLcd.Text) * 1.74532925199433E-02))
txtRefresh.Text = "1"
End If
End Sub
Function factorial(n As Double) As Double

' Debug.Print "Starting the calculation of " & n & " factorial"
    If n = 0 Then
        factorial = 1
    ElseIf n < 171 Then
' Debug.Print "Calling factorial(n) with n=" & n - 1
        factorial = factorial(n - 1) * n
    Else
     MsgBox "value should be less then 170"
    End If
' Debug.Print "Done calculating " & n & " factorial"

End Function

Private Sub cmdExp_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Exp(Val(frmCal.txtLcd.Text))
txtRefresh.Text = "1"
End If

End Sub

Private Sub cmdFact_Click()
If frmCal.cmdSwitch.Caption = "Off" Then
 If frmCal.txtLcd.Text = "" Then
  MsgBox "There is no value"
 Else
  frmCal.txtLcd.Text = factorial(frmCal.txtLcd.Text)
  txtRefresh.Text = "1"
 End If
End If
End Sub

Private Sub cmdHex_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
  If Val(frmCal.txtLcd.Text) < 1000000000 Then
   frmCal.txtLcd.Text = Hex(Val(frmCal.txtLcd.Text))
   txtRefresh.Text = "1"
   Else
  MsgBox "It can't be converted in Hexadecimal no."
 End If
End If

End Sub

Private Sub cmdLog_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Log(Val(frmCal.txtLcd.Text)) / 2.30258509299405)
txtRefresh.Text = "1"
End If

End Sub

Private Sub cmdLogn_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Log(Val(frmCal.txtLcd.Text)))
txtRefresh.Text = "1"
End If
End Sub

Private Sub cmdOct_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
 If Val(frmCal.txtLcd.Text) < 1000000000 Then
  frmCal.txtLcd.Text = Oct(Val(frmCal.txtLcd.Text))
  txtRefresh.Text = "1"
 Else
  MsgBox "It can't be converted in octal no."
 End If
End If

End Sub

Private Sub cmdRnd_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Rnd(Val(frmCal.txtLcd.Text))
txtRefresh.Text = "1"
End If

End Sub

Private Sub cmdRound_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Round(Val(frmCal.txtLcd.Text)))
txtRefresh.Text = "1"
End If

End Sub

Private Sub cmdSin_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Sin(Val(frmCal.txtLcd.Text) * 1.74532925199433E-02))
txtRefresh.Text = "1"
End If
End Sub

Private Sub cmdTan_Click()
Beep
If frmCal.cmdSwitch.Caption = "Off" Then
frmCal.txtLcd.Text = Str(Tan(Val(frmCal.txtLcd.Text) * 1.74532925199433E-02))
txtRefresh.Text = "1"
End If

End Sub

Private Sub Form_Load()
frmSci.Left = frmCal.Left + frmCal.Width
frmSci.Top = frmCal.Top + 1200
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmCal.cmdCon.Caption = "Scientific"
End Sub
