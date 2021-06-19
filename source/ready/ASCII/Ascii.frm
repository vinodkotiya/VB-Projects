VERSION 5.00
Begin VB.Form frmField 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VinASCII"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   FillColor       =   &H00FFFFFF&
   Icon            =   "Ascii.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   3975
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox text 
      Height          =   495
      Index           =   4
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox text 
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox text 
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox text 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Return Character"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtascii 
      Height          =   285
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox text 
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KEYCODE UP"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KEYCODE DOWN"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KEY ASCII"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Ascii Code"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command4_Click()
If Trim(txtascii.text) = "" Then Exit Sub
MsgBox "Ascii Code " & txtascii & " is for Character: " & Chr(Val(txtascii.text))
End Sub

Private Sub Form_Load()
On Error Resume Next
frmField.BackColor = Rnd(2) * 5000 * Second(Time)

Label1(0).ForeColor = frmField.BackColor / 256
Label1(1).ForeColor = frmField.BackColor / 256
Label1(2).ForeColor = frmField.BackColor / 256
Label1(3).ForeColor = frmField.BackColor / 256
text(1).text = "CursorX"
text(0).text = "CursorY"
'txtKey = "Ascii "
'txtDown = "Down"
'txtUp = "Up"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

text(1).text = Str(X)
text(0).text = Str(Y)
End Sub


Private Sub text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 3 Then text(3).text = Str(KeyCode)
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then text(2).text = Str(KeyAscii)
End Sub

Private Sub text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 4 Then text(4).text = Str(KeyCode)
End Sub

Private Sub txtascii_Change()
If IsNumeric(txtascii.text) = False Then
  MsgBox "Please Enter only Numeric Value"
  txtascii.text = ""
End If
End Sub


