VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DANGER GARDEN"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmStart.frx":1CCA
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgCredit 
      Height          =   525
      Left            =   360
      Picture         =   "frmStart.frx":3994
      Top             =   360
      Width           =   1200
   End
   Begin VB.Image ImgTank2 
      Height          =   720
      Left            =   8400
      MouseIcon       =   "frmStart.frx":5AA6
      MousePointer    =   99  'Custom
      Picture         =   "frmStart.frx":7770
      Top             =   3720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgTank1 
      Height          =   720
      Left            =   8400
      MouseIcon       =   "frmStart.frx":943A
      MousePointer    =   99  'Custom
      Picture         =   "frmStart.frx":B104
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgBil2 
      Height          =   720
      Left            =   6720
      Picture         =   "frmStart.frx":CDCE
      Top             =   3720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgBil1 
      Height          =   720
      Left            =   6720
      MouseIcon       =   "frmStart.frx":EA98
      MousePointer    =   99  'Custom
      Picture         =   "frmStart.frx":10762
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image ImgGh2 
      Height          =   720
      Left            =   5040
      MouseIcon       =   "frmStart.frx":1242C
      MousePointer    =   99  'Custom
      Picture         =   "frmStart.frx":140F6
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgGh1 
      Height          =   720
      Left            =   5040
      Picture         =   "frmStart.frx":14400
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image ImgKh2 
      Height          =   720
      Left            =   3360
      MouseIcon       =   "frmStart.frx":1470A
      MousePointer    =   99  'Custom
      Picture         =   "frmStart.frx":163D4
      Top             =   3720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image ImgKh1 
      Height          =   720
      Left            =   3360
      Picture         =   "frmStart.frx":1729E
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image ImgBg 
      Height          =   9000
      Left            =   0
      MouseIcon       =   "frmStart.frx":18168
      MousePointer    =   99  'Custom
      Picture         =   "frmStart.frx":19E32
      Top             =   -120
      Width           =   12000
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
'---------------------------------------------------------
'---------------- BY VINOD KOTIYA --------------------
'------ The code of this program is not very efficient
'------ because it was created on the early days of my
'------- visual basic computer programming.
'------- i made this programme without reading any VB book
'------- on the basis of my C++ experience i generally used
'------- if else statements
'------ code is easy and you can modify it
'------- in to a good code
'-------------------------------------------------------
'------ address S-2 shrimaya apartment sector-B/363
'------ sarvdharm colony bhopal
'---- fone +91-0755-2794428
'------ web: http://vinodkotiya.tripod.com     (without WWW)
'---- mail vinodkotiya24@rediffmail.com
'--------------------------------------------------------
'--------------------------------------------------------
Option Explicit

Private Sub imgCredit_Click()
Load frmCredit
frmCredit.Visible = True
End Sub

Private Sub imgTank1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgTank2.Visible = True
imgTank1.Visible = False
End Sub
Private Sub ImgTank2_Click()
Load frmastra
frmastra.Visible = True
Unload Me
Unload frmBball
Unload frmGhost
Unload frmBilli
End Sub
Private Sub ImgBil1_Click()
Load frmBilli
frmBilli.Visible = True
Unload Me
Unload frmBball
Unload frmGhost
Unload frmastra
End Sub

Private Sub ImgBg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgKh1.Visible = True
ImgKh2.Visible = False
ImgGh1.Visible = True
ImgGh2.Visible = False
imgBil2.Visible = True
ImgBil1.Visible = False
imgTank1.Visible = True
ImgTank2.Visible = False
End Sub

Private Sub imgBil2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgBil1.Visible = True
imgBil2.Visible = False
End Sub

Private Sub ImgGh1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgGh2.Visible = True
ImgGh1.Visible = False
End Sub

Private Sub ImgGh2_Click()
Load frmGhost
frmGhost.Visible = True
Unload Me
Unload frmBball
Unload frmBilli
Unload frmastra
End Sub

Private Sub ImgKh2_Click()
Load frmBball
frmBball.Visible = True
Unload Me
Unload frmGhost
Unload frmBilli
Unload frmastra
End Sub

Private Sub ImgKh1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgKh2.Visible = True
ImgKh1.Visible = False
End Sub
