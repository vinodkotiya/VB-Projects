VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   5160
   ClientTop       =   4230
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   1560
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'tranceparent
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
'Private Const LWA_ALPHA = &H3&
'H3 for control tranceparent else H2
Dim LWA_ALPHA As Long
Dim VIN As Byte
Dim vk As Integer
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40

Private Sub Form_Load()
On Error Resume Next
VIN = 0
'MsgBox Environ("OS")
If Environ("OS") = "Windows_NT" Then
' MsgBox "windows_nt"
 Timer1.Interval = Int(5000 / 255)
 LWA_ALPHA = &H2&
'''if os is not nt then do something else
Else
' MsgBox "Not windows nt"
frmSplash.Visible = True
  If Screen.Width > 15000 Then        'for 1024X 768
       SetWindowPos Me.hwnd, HWND_TOPMOST, 350, 270, 300, 200, SWP_SHOWWINDOW
  Else                                                    '800 X 600
       SetWindowPos Me.hwnd, HWND_TOPMOST, 260, 210, 300, 200, SWP_SHOWWINDOW
  End If
  'delay
  Dim i As Double
  i = Timer()
  While Timer() - i < 5  '3 second
   DoEvents
  Wend
  Load frmMain
    frmMain.Show
     Unload Me
End If

 
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
If VIN = 255 Then
   vk = vk + 1
   If vk = 128 Then
    Load frmMain
    frmMain.Show
     Unload Me
   End If
   Exit Sub
 End If
VIN = VIN + 1
Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)
Call GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Call SetLayeredWindowAttributes(Me.hwnd, 0, VIN, LWA_ALPHA)
If VIN = 1 Then
  frmSplash.Visible = True
  If Screen.Width > 15000 Then        'for 1024X 768
       SetWindowPos Me.hwnd, HWND_TOPMOST, 350, 270, 300, 200, SWP_SHOWWINDOW
  Else                                                    '800 X 600
       SetWindowPos Me.hwnd, HWND_TOPMOST, 260, 210, 300, 200, SWP_SHOWWINDOW
  End If
End If
End Sub
