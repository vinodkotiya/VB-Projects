VERSION 5.00
Begin VB.Form frmField 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   4350
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDown 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgRod 
      Height          =   150
      Left            =   4920
      Picture         =   "Rod1.frx":0000
      Top             =   7800
      Width           =   1050
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Form_Load()
frmField.BackColor = vbBlue
txtDown = "Down"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRod.Visible = True
If X < 11090 Then
imgRod.Left = X - 200
'imgRod.Left = imgRod.Left - Y
'Else If imgRod.Left < 0 Then
'imgRod.Visible = True
End If
End Sub

Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
txtDown.Text = Str(KeyCode)
If KeyCode = 37 Then
Beep
imgRod.Left = imgRod.Left - 200
ElseIf KeyCode = 39 Then
Beep
imgRod.Left = imgRod.Left + 200
End If

End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtUp_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

