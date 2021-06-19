VERSION 5.00
Begin VB.Form frmCredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDIT"
   ClientHeight    =   4335
   ClientLeft      =   1575
   ClientTop       =   2430
   ClientWidth     =   9030
   Icon            =   "frmCredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9030
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   0
      Picture         =   "frmCredit.frx":0ECA
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   597
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   1800
         Top             =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Web: http:\\vinodkotiya.tripod.com"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2550
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "vinodkotiya24@rediffmail.com"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   3720
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Dim BreakNow As Integer
Dim PenColor As Long
Dim mousex As Long
Dim mousey As Long
Dim speed As Integer  'falling speed


Private Sub DrawRoullette()
Dim R1 As Integer, R2 As Integer
Dim r As Integer
Dim pi As Double
On Error GoTo chupchap:
R1 = 1 'HScroll1.Value
R2 = 55 'HScroll2.Value - 80
'If R2 = 0 Then R2 = 10
r = 10 'HScroll4.Value
pi = 4 * Atn(1)

Dim loop1 As Integer, loop2 As Single
Dim t As Double, X As Double, Y As Double
Dim Rotations As Integer

If Int(R1 / R2) = R1 / R2 Then
    Rotations = 1
Else
    Rotations = Abs(R2 / 10)
    If Int(R2 / 10) <> R2 / 10 Then Rotations = 10 * Rotations
End If

For loop1 = 1 To Rotations
    PenColor = Picture1.Point(Picture1.ScaleWidth / 2 + X, Picture1.ScaleHeight / 2 + Y)
    For loop2 = 0 To 2 * pi Step pi / (4 * 360)
     
        t = loop1 * 2 * pi + loop2
        X = (R1 + R2) * Cos(t) - (R2 + r) * Cos(((R1 + R2) / R2) * t)
        Y = (R1 + R2) * Sin(t) - (R2 + r) * Sin(((R1 + R2) / R2) * t)
        Picture1.PSet (mousex + X, mousey + Y), PenColor
    Next
    DoEvents
    'Text1.Text = Str(loop1)
    BreakNow = True
    
    If loop1 = 30 Then
     Exit For
    End If
Next
Picture1.Refresh
 Exit Sub
chupchap:
End Sub



Private Sub cmdOk_Click()

End Sub


Private Sub Form_Load()
 counter = 0
 'mousex = frmCredit.Left + 200
 'mousey = frmCredit.Top + 100
 frmCredit.Top = 0 - frmCredit.Height
 speed = 170
 DrawRoullette
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
'Unload Me
End Sub

Private Sub Picture1_Click()
BreakNow = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BreakNow = True Then
mousex = X
mousey = Y
Picture1.Refresh
DrawRoullette
BreakNow = False

End If

End Sub

Private Sub Timer1_Timer()
On Error GoTo vinerror
If frmCredit.Top < 2895 Then
 frmCredit.Top = frmCredit.Top + speed
Else
 Timer1.Interval = 0
End If
If speed > 20 Then speed = speed - 2
Exit Sub
vinerror:
End Sub
