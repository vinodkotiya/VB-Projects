Attribute VB_Name = "SpiralMod"
Option Explicit
Global trignometry As Integer
Global PenColor As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const APP_NAME = "Eye strain"

Sub DrawRoullette()

Dim R1 As Integer, R2 As Integer
Dim r As Integer
Dim pi As Double
Form1.Picture1.DrawWidth = 1
PenColor = Rnd(4905673) * 65000
If PenColor < 10000 Then
 PenColor = vbGreen
End If
R1 = Rnd(354564) * 250 * Rnd(1)  'Form1.HScroll1.Value
If Rnd(45) < 0.2 Then
  R2 = Rnd(25) * (-50) 'Form1.HScroll2.Value - 80
Else
  R2 = Rnd(4566) * 50
End If
If R2 = 0 Then R2 = 10
r = Rnd(3) * 40  'Form1.HScroll4.Value
pi = 4 * Atn(1)

Dim loop1 As Integer, loop2 As Single
Dim t As Double, X As Double, Y As Double
Dim Rotations As Integer

If Int(R1 / R2) = R1 / R2 Then
    Rotations = Abs(Rnd(345) * 10) '1
Else
    Rotations = Abs(R2 / 10)
    If Int(R2 / 10) <> R2 / 10 Then Rotations = 10 * Rotations
End If

If trignometry = 1 Then
 For loop1 = 1 To Rotations
    
  On Error GoTo vinerror:
    For loop2 = 0 To 2 * pi Step pi / (4 * 360)
        t = loop1 * 2 * pi + loop2
        X = (R1 + R2) * Cos(t) - (R2 + r) * Cos(((R1 + R2) / R2) * t)
        Y = (R1 + R2) * Sin(t) - (R2 + r) * Sin(((R1 + R2) / R2) * t)
        Form1.Picture1.PSet (Form1.Picture1.ScaleWidth / 2 + X, Form1.Picture1.ScaleHeight / 2 + Y), PenColor
       
    Next
    DoEvents
 Next
ElseIf trignometry = 2 Then
For loop1 = 1 To Rotations
    
  On Error GoTo vinerror:
    For loop2 = 0 To 2 * pi Step pi / (4 * 360)
        t = loop1 * 2 * pi + loop2
        X = (R1 + R2) * Tan(t) - (R2 + r) * Cos(((R1 + R2) / R2) * t)
        Y = (R1 + R2) * Sin(t) - (R2 + r) * Sin(((R1 + R2) / R2) * t)
        Form1.Picture1.PSet (Form1.Picture1.ScaleWidth / 2 + X, Form1.Picture1.ScaleHeight / 2 + Y), PenColor
       
    Next
    DoEvents
 Next
ElseIf trignometry = 3 Then
 For loop1 = 1 To Rotations
    
  On Error GoTo vinerror:
    For loop2 = 0 To 2 * pi Step pi / (4 * 360)
        t = loop1 * 2 * pi + loop2
        X = (R1 + R2) * Cos(t) - (R2 + r) * Cos(((R1 + R2) / R2) * t)
        Y = (R1 + R2) * Tan(t) - (R2 + r) * Sin(((R1 + R2) / R2) * t)
        Form1.Picture1.PSet (Form1.Picture1.ScaleWidth / 2 + X, Form1.Picture1.ScaleHeight / 2 + Y), PenColor
       
    Next
    DoEvents
 Next
End If
Exit Sub
vinerror:
 MsgBox "An error has occured"
 End

End Sub

' See if another instance of the program is
' running in screen saver mode.
Public Sub CheckShouldRun()
    ' If no instance is running, we're safe.
    If Not App.PrevInstance Then Exit Sub

    ' See if there is a screen saver mode instance.
    If FindWindow(vbNullString, APP_NAME) Then End

    ' Set our caption so other instances can find
    ' us in the previous line.
    frmStart.Caption = APP_NAME
End Sub
