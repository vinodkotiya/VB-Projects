VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "FUNGAME"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fool"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdHitball 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hit the Ball"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBball 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bouncing Ball"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5790
      Left            =   0
      Picture         =   "frmStart.frx":0ECA
      Top             =   0
      Width           =   8385
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBball_Click()
Load frmBball
frmBball.Visible = True
 Me.Hide

End Sub

Private Sub cmdHitball_Click()
Load frmField
frmField.Visible = True
Me.Hide

End Sub

Private Sub Command1_Click()
Load frmFool
frmFool.Visible = True
Me.Hide
End Sub
