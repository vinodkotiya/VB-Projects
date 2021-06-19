VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Transparent form by vinod kotiya"
   ClientHeight    =   8070
   ClientLeft      =   3240
   ClientTop       =   1605
   ClientWidth     =   7020
   Icon            =   "transparent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "transparent.frx":030A
   ScaleHeight     =   8070
   ScaleWidth      =   7020
   Visible         =   0   'False
   Begin VB.CommandButton cmdSource 
      Caption         =   "E"
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   20
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "*"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   19
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "*"
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   18
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "S"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   17
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "O"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   16
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "U"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   15
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "R"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "C"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   13
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "*"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   7320
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1095
      LargeChange     =   400
      Left            =   5520
      Max             =   11000
      Min             =   1
      SmallChange     =   100
      TabIndex        =   11
      Top             =   480
      Value           =   1
      Width           =   135
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   500
      Left            =   4920
      Max             =   8000
      Min             =   1
      SmallChange     =   100
      TabIndex        =   10
      Top             =   960
      Value           =   1
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   3840
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transeparency"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Controls"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Form"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4920
      Top             =   4680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Form Color"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   6360
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "255"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Text            =   "100 %"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   2520
      Max             =   255
      TabIndex        =   1
      Top             =   3120
      Value           =   255
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Semi Transparent"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   7
      Left            =   -360
      Shape           =   2  'Oval
      Top             =   7440
      Width           =   855
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   6
      Left            =   6360
      Shape           =   2  'Oval
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   5
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   4
      Left            =   4200
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   975
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6480
      Picture         =   "transparent.frx":CC9C
      Top             =   1320
      Width           =   480
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   3840
      Shape           =   2  'Oval
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   100
      Index           =   2
      X1              =   -480
      X2              =   240
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderWidth     =   100
      Index           =   1
      X1              =   -840
      X2              =   120
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderWidth     =   100
      Index           =   0
      X1              =   6120
      X2              =   6840
      Y1              =   3360
      Y2              =   3480
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   1
      Left            =   -480
      Shape           =   2  'Oval
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   0
      Left            =   6360
      Shape           =   2  'Oval
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   720
      Shape           =   2  'Oval
      Top             =   6960
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   6120
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Opacity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***      *** ***   *****  ***   *******    *******
'  ***    ***  ***   *****  ***  ***   ***   ***  ****
'   ***  ***   ***   *** ** ***  ***   ***   ***   ****
'    ******    ***   ***  *****  ***   ***   ***  ****
'     ****     ***   ***   ****   *******    *******
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Programmer : VINOD KOTIYA
'  B.E. (Information Technology)
'  Semester V
'  University Institute of Technology
'  Rajeev Gandhi Prodyogiki Vishwavidyalaya Bhopal.
'  Address: S-2 ShreeMaya Apartment Sector-B/363
'           Sarvdharm Colony Bhopal-42 (India)
'  Email: vinodkotiya24@rediffmail.com
'  Web : http://vinodkotiya.tripod.com
'Get full project source code of games and other software
' at my website or contact by mail.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
'Private Const LWA_ALPHA = &H3&
'H3 for control tranceparent else H2
Dim LWA_ALPHA As Long
Dim VIN As Byte




Private Sub cmdSource_Click(Index As Integer)
Load frmSource
frmSource.Visible = True
End Sub

Private Sub Command1_Click()
HScroll1.Value = 128
End Sub

Private Sub Command2_Click()
Form1.BackColor = 65000 * Rnd(Second(Time) * Minute(Time))
End Sub


Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
If Environ("OS") = "Windows_NT" Then
Dim displayTime As Integer
displayTime = 1000   'in milliseconds
VIN = 0

  Timer1.Interval = Int(displayTime / 255)

LWA_ALPHA = &H2&
Else
 Me.Show
 MsgBox "Tranceparancy Feature will only work on windowsXP/NT" & vbCrLf & _
  "Your OS not support the required API"
End If
End Sub



Private Sub HScroll1_Change()
On Error GoTo transError
Dim Level As Byte
Level = HScroll1.Value
Text1.Text = (Level / 255) * 100 & "  %"
Text2.Text = Level
Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)
Call GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Call SetLayeredWindowAttributes(Me.hwnd, 0, Level, LWA_ALPHA)
Exit Sub
transError:
MsgBox "Tranceparancy Feature will only work on windowsXP/NT" & vbCrLf & _
  "Your OS not support the required API"
End Sub

Private Sub HScroll2_Change()
Form1.Left = HScroll2.Value
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
 LWA_ALPHA = &H2&
 Label1.Visible = True
 HScroll1.Visible = True
 Text1.Visible = True
 Text2.Visible = True
Else
 LWA_ALPHA = &H3&
 Label1.Visible = False
 HScroll1.Visible = False
 Text1.Visible = False
 Text2.Visible = False
End If
HScroll1_Change
End Sub

Private Sub Timer1_Timer()
If VIN = 255 Then
  Timer1.Interval = 0
   Exit Sub
 End If
VIN = VIN + 1
Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)
Call GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Call SetLayeredWindowAttributes(Me.hwnd, 0, VIN, LWA_ALPHA)
If VIN = 1 Then Form1.Visible = True
End Sub

Private Sub VScroll1_Change()
Form1.Top = VScroll1.Value
End Sub
