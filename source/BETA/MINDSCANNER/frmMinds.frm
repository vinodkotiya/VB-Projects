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
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACTRESSES"
      Height          =   495
      Left            =   4440
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdActor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACTORS"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Image imgCredit2 
      Height          =   525
      Left            =   3120
      Picture         =   "frmMinds.frx":030A
      Top             =   4260
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgCredit1 
      Height          =   525
      Left            =   3120
      Picture         =   "frmMinds.frx":06DA
      Top             =   4260
      Width           =   1200
   End
   Begin VB.Image imgMindScan 
      Height          =   795
      Left            =   1200
      Picture         =   "frmMinds.frx":0BDA
      Top             =   4800
      Width           =   5160
   End
   Begin VB.Image imgTaj 
      Height          =   5625
      Left            =   0
      Picture         =   "frmMinds.frx":246D
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCredit2.Visible = False
imgCredit1.Visible = True

End Sub

Private Sub imgCredit2_Click()
frmCredit.Visible = True
End Sub

Private Sub imgCredit1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCredit1.Visible = False
imgCredit2.Visible = True
End Sub

Private Sub imgMindScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCredit2.Visible = False
imgCredit1.Visible = True

End Sub

Private Sub imgTaj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCredit2.Visible = False
imgCredit1.Visible = True

End Sub
