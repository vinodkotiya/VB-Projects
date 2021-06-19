VERSION 5.00
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VINCALCULATOR"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdPer 
      Caption         =   "%"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdNul 
      Caption         =   "0"
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdFor 
      Caption         =   "4"
      Height          =   495
      Left            =   960
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdFiv 
      Caption         =   "5"
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdSix 
      Caption         =   "6"
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdMul 
      Caption         =   "X"
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "_"
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdPlu 
      Caption         =   "+"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdNin 
      Caption         =   "9"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdAt 
      Caption         =   "8"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdSev 
      Caption         =   "7"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdThr 
      Caption         =   "3"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "2"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "1"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdEqu 
      Caption         =   "="
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdDec 
      Caption         =   "."
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdDiv 
      Caption         =   "/"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdCredit 
      Caption         =   "Credit"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "OFF"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtLcd 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmCal.frx":0000
      Top             =   -480
      Width           =   7500
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
Dim Flag As Integer       'used for taking Data2
Dim Ref As Integer       'refreshing LCD after =



Private Sub cmdAt_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "8"
End Sub

Private Sub cmdCredit_Click()
frmCredit.Visible = True
frmCal.Visible = False
End Sub

Private Sub cmdDec_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "."
End Sub

Private Sub cmdDiv_Click()
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 4
End If
End Sub

Private Sub cmdPer_Click()
If Flag = 1 Then
Data2 = Val(txtLcd.Text)
Flag = 0
End If
'Op = 5
Res = (Data1 * Data2) / 100
txtLcd.Text = Str(Res)
Ref = 1
End Sub

Private Sub cmdPlu_Click()
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 1
End If
End Sub


Private Sub cmdEqu_Click()
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
End Sub

Private Sub cmdFiv_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "5"
End Sub

Private Sub cmdFor_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "4"
End Sub

Private Sub cmdMin_Click()
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 2
End If
End Sub

Private Sub cmdMul_Click()
If Flag = 0 Then
Data1 = Val(txtLcd.Text)
Flag = 1
txtLcd.Text = ""
Op = 3
End If

End Sub

Private Sub cmdNin_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "9"
End Sub

Private Sub cmdNul_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "0"
End Sub

Private Sub cmdOff_Click()
End
End Sub

Private Sub cmdOne_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "1"
End Sub


Private Sub cmdSev_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "7"
End Sub

Private Sub cmdSix_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "6"
End Sub

Private Sub cmdThr_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "3"
End Sub

Private Sub cmdTwo_Click()
If Ref = 1 Then
txtLcd.Text = ""
Ref = 0
End If
txtLcd.Text = txtLcd.Text & "2"
End Sub

Private Sub Form_Load()
Flag = 0
Ref = 0
End Sub

Private Sub txtLcd_KeyDown(KeyCode As Integer, Shift As Integer)
txtTemp.Text = Str(KeyCode)
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

