VERSION 5.00
Begin VB.Form frmButton 
   BorderStyle     =   0  'None
   ClientHeight    =   675
   ClientLeft      =   390
   ClientTop       =   525
   ClientWidth     =   9780
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmButton.frx":0000
   ScaleHeight     =   675
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Height          =   375
      Index           =   1
      Left            =   9000
      Picture         =   "frmButton.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdHelp 
      Height          =   375
      Index           =   0
      Left            =   8520
      Picture         =   "frmButton.frx":1DB4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdProj 
      Height          =   495
      Index           =   1
      Left            =   600
      Picture         =   "frmButton.frx":213E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open Project"
      Top             =   90
      Width           =   495
   End
   Begin VB.CommandButton cmdProj 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmButton.frx":24C8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "New Project"
      Top             =   90
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Index           =   1
      Left            =   1680
      Picture         =   "frmButton.frx":390A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save Project As..."
      Top             =   90
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Index           =   0
      Left            =   1200
      Picture         =   "frmButton.frx":3C94
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save Project"
      Top             =   90
      Width           =   495
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      Height          =   495
      Left            =   6480
      Picture         =   "frmButton.frx":401E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Build"
      Top             =   90
      Width           =   615
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   495
      Left            =   5640
      Picture         =   "frmButton.frx":45A8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Compile"
      Top             =   90
      Width           =   735
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   7320
      Picture         =   "frmButton.frx":4B32
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Run"
      Top             =   90
      Width           =   615
   End
   Begin VB.Image imgStepNo 
      Height          =   480
      Index           =   4
      Left            =   4920
      Picture         =   "frmButton.frx":4EBC
      Top             =   90
      Width           =   480
   End
   Begin VB.Image imgStepNo 
      Height          =   480
      Index           =   3
      Left            =   4320
      Picture         =   "frmButton.frx":5605
      Top             =   90
      Width           =   480
   End
   Begin VB.Image imgStepNo 
      Height          =   480
      Index           =   2
      Left            =   3720
      Picture         =   "frmButton.frx":5D36
      Top             =   90
      Width           =   480
   End
   Begin VB.Image imgStepNo 
      Height          =   480
      Index           =   1
      Left            =   3120
      Picture         =   "frmButton.frx":6485
      Top             =   90
      Width           =   480
   End
   Begin VB.Image imgStepOn 
      Height          =   480
      Index           =   4
      Left            =   4920
      Picture         =   "frmButton.frx":6BC8
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOn 
      Height          =   480
      Index           =   3
      Left            =   4320
      Picture         =   "frmButton.frx":72A2
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOn 
      Height          =   480
      Index           =   2
      Left            =   3720
      Picture         =   "frmButton.frx":7984
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOn 
      Height          =   480
      Index           =   1
      Left            =   3120
      Picture         =   "frmButton.frx":7D28
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOver 
      Height          =   480
      Index           =   4
      Left            =   4920
      Picture         =   "frmButton.frx":8402
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOver 
      Height          =   480
      Index           =   3
      Left            =   4320
      Picture         =   "frmButton.frx":87A1
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOver 
      Height          =   480
      Index           =   2
      Left            =   3720
      Picture         =   "frmButton.frx":8CDC
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOver 
      Height          =   480
      Index           =   1
      Left            =   3120
      Picture         =   "frmButton.frx":921A
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOn 
      Height          =   480
      Index           =   0
      Left            =   2520
      Picture         =   "frmButton.frx":95B9
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepOver 
      Height          =   480
      Index           =   0
      Left            =   2520
      Picture         =   "frmButton.frx":9C78
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStepNo 
      Height          =   480
      Index           =   0
      Left            =   2520
      Picture         =   "frmButton.frx":A1A7
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim onlyonce As Boolean

Private Sub cmdBuild_Click()
mdifrmMain.mnuBuild_Click
End Sub

Private Sub cmdCompile_Click()

mdifrmMain.mnuCompile_Click

End Sub



Private Sub cmdHelp_Click(Index As Integer)
If Index = 0 Then
 mdifrmMain.mnuHelp_Click (0)
Else
 mdifrmMain.mnuHelp_Click (2)
End If
End Sub

Private Sub cmdProj_Click(Index As Integer)
If Index = 0 Then
 mdifrmMain.mnuNew_Click
Else
 mdifrmMain.mnuOpen_Click
End If
End Sub

Private Sub cmdRun_Click()
mdifrmMain.mnuRun_Click
End Sub

Private Sub cmdSave_Click(Index As Integer)
mdifrmMain.mnuSave_Click (Index)
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Index As Integer
If onlyonce Then
For Index = 0 To 4
If imgStepOn(Index).Visible = False Then
 imgStepNo(Index).Visible = True
 imgStepOver(Index).Visible = False
End If
Next
onlyonce = False
End If
End Sub

Private Sub imgStepNo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStepNo(Index).Visible = False
imgStepOver(Index).Visible = True
onlyonce = True
End Sub

Public Sub imgStepOver_Click(Index As Integer)
Dim i As Integer
imgStepNo(Index).Visible = False
imgStepOver(Index).Visible = False
imgStepOn(Index).Visible = True
For i = 0 To 4
 If i <> Index Then
  imgStepOn(i).Visible = False
  imgStepNo(i).Visible = True
 End If
Next
mdifrmMain.mnuWiz_Click (Index)
End Sub

'Public Sub optStep_Click(Index As Integer)
'mdifrmMain.mnuWiz_Click (Index)
'optStep(Index).Value = True'

'End Sub
