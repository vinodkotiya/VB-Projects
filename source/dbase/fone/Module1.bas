Attribute VB_Name = "Module1"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const GWL_EXSTYLE = -20
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2&
'H3 for control tranceparent else H2

