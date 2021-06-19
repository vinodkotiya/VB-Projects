Attribute VB_Name = "Module1"
Option Explicit

 Sub Main()
  If 12 < Hour(Time()) And Hour(Time()) < 5 Then
    MsgBox "exit"
    Else
    MsgBox "yo" & Time()
  End If
End Sub
