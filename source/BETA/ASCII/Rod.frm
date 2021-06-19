VERSION 5.00
Begin VB.Form Form1 
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
   Begin VB.Image imgRod 
      Height          =   1500
      Left            =   0
      MousePointer    =   99  'Custom
      Picture         =   "Rod.frx":0000
      Top             =   6960
      Width           =   12000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRod.Visible = True

End Sub

Private Sub imgRod_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button > 0 Then
imgRod.Left = imgRod.Left + X + 4700
imgRod.Left = imgRod.Left - Y - 9000
End If
End Sub
