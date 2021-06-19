VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rubber Lines Demo"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XStart, YStart As Single
Dim XOld, YOld As Single

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    XStart = X
    YStart = Y
    XOld = XStart
    YOld = YStart
    Form1.DrawMode = 7
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Form1.Line (XStart, YStart)-(XOld, YOld)
    Form1.Line (XStart, YStart)-(X, Y)
    XOld = X
    YOld = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Form1.DrawMode = 13
    Form1.Line (XStart, YStart)-(XOld, YOld)
    Form1.Line (XStart, YStart)-(X, Y)
End Sub
