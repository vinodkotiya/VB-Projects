VERSION 5.00
Begin VB.Form frmMinds 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIND-SCANNER"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   FillColor       =   &H00FF80FF&
   FillStyle       =   0  'Solid
   Icon            =   "frmMinds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdActress 
      Caption         =   "ACTRESSES"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdActor 
      Caption         =   "ACTORS"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   1200
      Picture         =   "frmMinds.frx":030A
      Top             =   4320
      Width           =   5160
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "frmMinds.frx":1B9D
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmMinds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()

End Sub

Private Sub cmdActor_Click()
Load frmActress
frmActors.Visible = True
Unload Me
End Sub

Private Sub cmdActress_Click()
Load frmActress
frmActress.Visible = True
Unload Me
End Sub
