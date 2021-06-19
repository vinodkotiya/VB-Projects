VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "v/s Player"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "v/s Computer"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton opt3 
      Caption         =   "Fastest"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Faster"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Fast"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame fraLevel 
      Caption         =   "LEVEL"
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmField.Com = 1
Load frmField
frmField .Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
frmField.Com = 0
Load frmField
frmField .Visible = True
Unload Me
End Sub

Private Sub Form_Load()
opt2.Value = True
End Sub

Private Sub opt1_Click()
frmField.Level = 20
End Sub

Private Sub opt2_Click()
frmField.Level = 40
End Sub

Private Sub opt3_Click()
frmField.Level = 60
End Sub
