VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALVIN"
   ClientHeight    =   6135
   ClientLeft      =   705
   ClientTop       =   2355
   ClientWidth     =   7155
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   Icon            =   "calvinsimple1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7155
   Begin VB.CommandButton cmdOperator 
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
      Index           =   4
      Left            =   6360
      Picture         =   "calvinsimple1.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
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
      Index           =   3
      Left            =   6360
      Picture         =   "calvinsimple1.frx":2478
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
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
      Index           =   2
      Left            =   6360
      Picture         =   "calvinsimple1.frx":2C26
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3720
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   9
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   8
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   7
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   6
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   5
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   4
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   3
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   2
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   80
      Width           =   1095
   End
   Begin VB.CommandButton cmdCredit 
      BackColor       =   &H00FFFF80&
      Caption         =   "CREDIT"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto Copy Result"
      Height          =   495
      Left            =   3960
      MaskColor       =   &H008080FF&
      TabIndex        =   22
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Set on Top of All Windows"
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdOperator 
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
      Index           =   1
      Left            =   6360
      Picture         =   "calvinsimple1.frx":2F59
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdCon 
      BackColor       =   &H0000FFFF&
      Caption         =   "Scientific"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdXpy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "x^y"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdMr 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MR"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdMp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "M+"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdCe 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CE"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdReci 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1/x"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   615
   End
   Begin VB.PictureBox picF6 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      Picture         =   "calvinsimple1.frx":3707
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "--->"
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdSqrt 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Sqrt."
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Left            =   2760
      Picture         =   "calvinsimple1.frx":3A11
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   615
   End
   Begin VB.Timer timFace 
      Interval        =   200
      Left            =   5640
      Top             =   720
   End
   Begin VB.PictureBox picF5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2760
      Picture         =   "calvinsimple1.frx":3AD5
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
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
      Left            =   2760
      Picture         =   "calvinsimple1.frx":3DDF
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
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
      Left            =   2760
      Picture         =   "calvinsimple1.frx":40E9
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
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
      Left            =   2760
      Picture         =   "calvinsimple1.frx":43F3
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
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
      Left            =   2760
      Picture         =   "calvinsimple1.frx":46FD
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPer 
      BackColor       =   &H00C0FFC0&
      Caption         =   "%"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdDigit 
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
      Index           =   0
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FF00FF&
      Caption         =   "On"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtLcd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "calvinsimple1.frx":4A07
      Top             =   120
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5655
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9975
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"calvinsimple1.frx":4A43
   End
   Begin VB.Image imgScroll 
      Height          =   450
      Left            =   -4680
      Picture         =   "calvinsimple1.frx":4ABC
      Top             =   5760
      Width           =   12000
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   2640
      Picture         =   "calvinsimple1.frx":1643E
      Top             =   0
      Width           =   4500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   435
      Left            =   0
      Top             =   5700
      Width           =   7935
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data1 As Double
Dim Data2 As Double
Dim Res As Double
Dim Operator As Integer
Dim isoperatorclick As Boolean 'IF ANY OPERATOR IS PRESSED SET TRUE
Dim refreshscreen As Boolean  'when true refresh the screen


Dim F As Integer  'display note
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
Data1 = 0
Data2 = 0
RichTextBox1.Text = RichTextBox1.Text & "   " & Str(Res) & Chr(13) & "    "
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
Beep
If cmdSwitch.Caption = "Off" Then
If InStr(txtLcd.Text, ".") = False Then
  If txtLcd.Text = "" Then
   txtLcd.Text = "0."
  Else
   txtLcd.Text = txtLcd.Text & "."
   End If
 Else
 MsgBox "Don't you know that only one decimal should be placed in a number"
End If
End If
End Sub

Private Sub cmdDigit_Click(Index As Integer)
If cmdSwitch.Caption = "Off" Then
 RichTextBox1.Text = RichTextBox1.Text + cmdDigit(Index).Caption
 If refreshscreen = True Then
  txtLcd.Text = ""
  txtLcd.Text = txtLcd.Text + cmdDigit(Index).Caption
  refreshscreen = False
 Else
  txtLcd.Text = txtLcd.Text + cmdDigit(Index).Caption
 End If
End If
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

Private Sub cmdOperator_Click(Index As Integer)
On Error GoTo vinerror
If Trim(txtLcd.Text) <> " " Then
 refreshscreen = True 'now refresh the screen when pree digits

 If isoperatorclick = False Then
'  Operator = Index
  Data1 = txtLcd.Text
  isoperatorclick = True  'make again true b/c cmdequ make it false
  
 Else
 
 'Data1 = txtLcd.Text
Operator = Index
 'isoperatorclick = True

  RichTextBox1.Text = RichTextBox1.Text + Chr(13)
  If Index = 1 Then
   Data1 = Data1 + Val(txtLcd.Text)
   RichTextBox1.Text = RichTextBox1.Text + " +   "
  ElseIf Index = 2 Then
   Data1 = Data1 - Val(txtLcd.Text)
   RichTextBox1.Text = RichTextBox1.Text + " -   "
  ElseIf Index = 3 Then
  Data1 = Data1 * Val(txtLcd.Text)
   RichTextBox1.Text = RichTextBox1.Text + " X   "
  ElseIf Index = 4 Then
  Data1 = Data1 / Val(txtLcd.Text)
   RichTextBox1.Text = RichTextBox1.Text + " /   "
  End If
  txtLcd.Text = Str(Data1)
 End If
End If
 
 Exit Sub
 
vinerror:
 MsgBox "Please first turn on the calculator"
 
End Sub

Private Sub cmdPer_Click()
Beep
'If Flag = 1 Then
Data2 = Val(txtLcd.Text)
'Flag = 0
'End If
Res = (Data1 * Data2) / 100
txtLcd.Text = Str(Res)
'Ref = 1
If cmdSwitch.Caption = "On" Then
 If timFace.Interval < 11 Then
  timFace.Interval = 20
 Else
  timFace.Interval = timFace.Interval - 10
 End If
End If
End Sub



Private Sub cmdEqu_Click()
RichTextBox1.Text = RichTextBox1.Text & Chr(13) & "-------------" & Chr(13)
Beep
Data2 = Val(txtLcd.Text)
MsgBox Data1 & Data2
If Operator = 1 Then
Res = Data1 + Data2
ElseIf Operator = 2 Then
Res = Data1 - Data2
ElseIf Operator = 3 Then
Res = Data1 * Data2
ElseIf Operator = 4 Then
  If Data2 > 0 Then
  Res = Data1 / Data2
  ElseIf Data2 < 0 Then
  Res = Data1 / Data2
  Else
  txtLcd.Text = " CAN NOT DIVIDE BY ZERO U FOOL"
  End If
ElseIf Operator = 5 Then
Res = Data1 ^ Data2
End If
'If (Res > (4.94065645841247E-324) And Res < (1.7E+308)) Or (Res > (-4.94065645841247E-324) And Res < (-1.7E+308)) Then
txtLcd.Text = Str(Res)
RichTextBox1.Text = RichTextBox1.Text & "   " & Str(Res) & Chr(13) & "    "
'End If
isoperatorclick = False
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
  txtLcd.Text = "0" & Str(Res)
  isoperatorclick = True
 End If
End If

End Sub

Private Sub cmdSqrt_Click()
Beep
Res = Sqr(Val(txtLcd.Text))
txtLcd.Text = Str(Res)
isoperatorclick = True

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
txtLcd.SetFocus
isoperatorclick = False
refreshscreen = True 'when true refresh the screen
RichTextBox1.Text = RichTextBox1.Text & "   " & Str(Res) & Chr(13) & "    "
End If
End Sub









Private Sub cmdXpy_Click()
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Operator = 5
End If

End Sub

Private Sub Command1_Click()
Dim pos As Integer

  pos = InStr(txtLcd.Text, "+")
  txtLcd.Text = Left(txtLcd.Text, pos - 1)
   MsgBox pos
End Sub

Private Sub Form_Load()


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
txtDate.Text = Now
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
isoperatorclick = True
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
'Dim pos As Integer
'If isoperatorclick = True Then
' txtLcd.Text = ""   'if any operator is pressed then clear screen for newdata
 'isoperatorclick = False
'End If

'If KeyCode = 187 Then        '+WITH SHIFT
' If Shift And vbShiftMask Then
'  cmdOperator_Click (1)
  'pos = InStr(txtLcd.Text, "+")
 ' txtLcd.Text = Right(txtLcd.Text, pos + 1)
  'MsgBox pos
 'Else
  'cmdEqu_Click
 '  End If
  
'ElseIf KeyCode = 189 Or KeyCode = 109 Then    'MINUS
 ' cmdOperator_Click (2)
'ElseIf KeyCode = 56 Then    '*WITH SHIFT
 'If Shift And vbShiftMask Then
  'cmdOperator_Click (3)
 'End If
'ElseIf KeyCode = 191 Or KeyCode = 111 Then    'DIVISION
  'cmdOperator_Click (4)

  

'End If

End Sub

Private Sub txtLcd_KeyUp(KeyCode As Integer, Shift As Integer)
Dim pos As Integer
On Error GoTo vinerror
If Trim(txtLcd.Text) <> " " Then
 If KeyCode = 187 Then        '+WITH SHIFT
  If Shift And vbShiftMask Then
   pos = InStr(txtLcd.Text, "+")
   txtLcd.Text = Left(txtLcd.Text, pos - 1)
   cmdOperator_Click (1)
  Else
    pos = InStr(txtLcd.Text, "=")
    If pos > 0 Then
     txtLcd.Text = Left(txtLcd.Text, pos - 1)
    End If
    cmdEqu_Click
   End If
 ElseIf KeyCode = 189 Or KeyCode = 109 Then    'MINUS
        pos = InStr(txtLcd.Text, "-")
   txtLcd.Text = Left(txtLcd.Text, pos - 1)
   cmdOperator_Click (2)
 ElseIf KeyCode = 56 Then    '*WITH SHIFT
  If Shift And vbShiftMask Then
     pos = InStr(txtLcd.Text, "*")
    If pos > 0 Then    'if 8 is pressed
     txtLcd.Text = Left(txtLcd.Text, pos - 1)
    End If
    cmdOperator_Click (3)
  End If
 ElseIf KeyCode = 191 Or KeyCode = 111 Then    'DIVISION
     pos = InStr(txtLcd.Text, "/")
   txtLcd.Text = Left(txtLcd.Text, pos - 1)
    cmdOperator_Click (4)
  End If
End If 'end of txtLcd.Text = ""
  
If isoperatorclick = True And (KeyCode > 47 And KeyCode < 58) Then
 txtLcd.Text = Chr(KeyCode)   'if any operator is pressed then clear screen for newdata
 isoperatorclick = False
End If
  
  
  Exit Sub
vinerror:
 MsgBox "Invalid values are entered"
 txtLcd.Text = ""
End Sub
