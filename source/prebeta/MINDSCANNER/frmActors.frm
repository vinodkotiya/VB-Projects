VERSION 5.00
Begin VB.Form frmActors 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdResult 
      Caption         =   "RESULT"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "START"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgCredit2 
      Height          =   330
      Left            =   8280
      Picture         =   "frmActors.frx":0000
      Top             =   480
      Width           =   750
   End
   Begin VB.Image imgCredit1 
      Height          =   330
      Left            =   8280
      Picture         =   "frmActors.frx":058D
      Top             =   480
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgComs 
      Height          =   540
      Left            =   1920
      Picture         =   "frmActors.frx":0A6E
      Top             =   3120
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.Image imgCom 
      Height          =   2430
      Left            =   240
      Picture         =   "frmActors.frx":31BF
      Top             =   2400
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgGuide 
      Height          =   1080
      Left            =   120
      Picture         =   "frmActors.frx":4A5B
      Top             =   1250
      Width           =   7605
   End
   Begin VB.Image imgClose2 
      Height          =   300
      Left            =   8740
      Picture         =   "frmActors.frx":52D6
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose1 
      Height          =   300
      Left            =   8745
      Picture         =   "frmActors.frx":5818
      Top             =   90
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   120
      Picture         =   "frmActors.frx":5D5A
      Top             =   480
      Width           =   5160
   End
   Begin VB.Image imgChoice 
      Height          =   300
      Left            =   120
      Picture         =   "frmActors.frx":74CA
      Top             =   1920
      Visible         =   0   'False
      Width           =   8850
   End
   Begin VB.Image imgBobby 
      Height          =   1230
      Left            =   7680
      Picture         =   "frmActors.frx":A74C
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Image imgSmk 
      Height          =   1230
      Left            =   3120
      Picture         =   "frmActors.frx":C156
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Image imgAnil 
      Height          =   1230
      Left            =   240
      Picture         =   "frmActors.frx":DB60
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Image imgAjay 
      Height          =   1230
      Left            =   1560
      Picture         =   "frmActors.frx":F56A
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Image imgSrk 
      Height          =   1230
      Left            =   5640
      Picture         =   "frmActors.frx":10F74
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Image imgAkki 
      Height          =   1230
      Left            =   6240
      Picture         =   "frmActors.frx":1297E
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Image imgRithik 
      Height          =   1230
      Left            =   7800
      Picture         =   "frmActors.frx":14388
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Image imgAkshay 
      Height          =   1230
      Left            =   2640
      Picture         =   "frmActors.frx":15D92
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Image imgAb 
      Height          =   1230
      Left            =   240
      Picture         =   "frmActors.frx":1779C
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Image imgAmir 
      Height          =   1230
      Left            =   4560
      Picture         =   "frmActors.frx":191A6
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Image imgBg 
      Height          =   4980
      Left            =   0
      Picture         =   "frmActors.frx":1ABB0
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmActors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Scan As Integer, Start As Integer


Private Sub cmdNo_Click()
If cmdList.Caption = "List1" Then
 Scan = Scan + 0
imgSrk.Visible = False
imgRithik.Visible = True
imgAkshay.Visible = True
imgAb.Visible = True
imgAjay.Visible = False
imgAmir.Visible = True
imgSmk.Visible = False
imgBobby.Visible = False
imgAkki.Visible = True
imgAnil.Visible = False
imgAkshay.Top = imgAkshay.Top + 1230
imgAmir.Top = imgAmir.Top - 1300
cmdList.Caption = "List2"
ElseIf cmdList.Caption = "List2" Then
 Scan = Scan + 0
imgSrk.Visible = True
imgRithik.Visible = False
imgAkshay.Visible = False
imgAb.Visible = True
imgAjay.Visible = False
imgAmir.Visible = True
imgSmk.Visible = True
imgBobby.Visible = False
imgAkki.Visible = True
imgAnil.Visible = True
cmdList.Caption = "List3"
ElseIf cmdList.Caption = "List3" Then
 Scan = Scan + 0
imgSrk.Visible = False
imgRithik.Visible = False
imgAkshay.Visible = True
imgAb.Visible = False
imgAjay.Visible = True
imgAmir.Visible = True
imgSmk.Visible = False
imgBobby.Visible = True
imgAkki.Visible = False
imgAnil.Visible = True
imgAb.Left = 7800
cmdList.Caption = "List4"
'ElseIf cmdList.Caption = "List4" Then
'Scan = Scan + 0
'imgSrk.Visible = True
'imgRithik.Visible = True
'imgAkshay.Visible = True
'imgAb.Visible = True
'imgAjay.Visible = False
'imgAmir.Visible = False
'imgSmk.Visible = True
'imgBobby.Visible = True
'imgAkki.Visible = False
'imgAnil.Visible = False
'cmdList.Caption = "List5"

ElseIf cmdList.Caption = "List4" Then
 'Scan = Scan + 0
imgSrk.Visible = False
imgRithik.Visible = False
imgAkshay.Visible = False
imgAb.Visible = False
imgAjay.Visible = False
imgAmir.Visible = False
imgSmk.Visible = False
imgBobby.Visible = False
imgAkki.Visible = False
imgAnil.Visible = False
cmdList.Visible = False
imgChoice.Visible = False
cmdYes.Visible = False
cmdNo.Visible = False
cmdResult.Visible = True

End If
End Sub


Private Sub cmdRestart_Click()
If Start = 0 Then
 If cmdList.Caption = "List1" Then
 imgSrk.Visible = True
 imgRithik.Visible = True
 imgAkshay.Visible = True
 imgAb.Visible = True
 imgAjay.Visible = True
 imgAmir.Visible = True
 imgSmk.Visible = False
 imgBobby.Visible = False
 imgAkki.Visible = False
 imgAnil.Visible = False
 imgChoice.Visible = True
 cmdYes.Visible = True
 cmdNo.Visible = True
 cmdList.Visible = True
 cmdRestart.Visible = False
 imgGuide.Visible = False
 Start = 1
 End If
'End If
ElseIf Start = 1 Then
 frmActors.Visible = False
 Load frmMinds
 frmMinds.Visible = True
 Unload frmActress
 Unload frmActors
 End If
End Sub

Private Sub cmdYes_Click()
If cmdList.Caption = "List1" Then
 Scan = Scan + 3           'column 3
imgSrk.Visible = False
imgRithik.Visible = True
imgAkshay.Visible = True
imgAb.Visible = True
imgAjay.Visible = False
imgAmir.Visible = True
imgSmk.Visible = False
imgBobby.Visible = False
imgAkki.Visible = True
imgAnil.Visible = False
cmdList.Caption = "List2"
ElseIf cmdList.Caption = "List2" Then
 Scan = Scan + 4
imgSrk.Visible = True
imgRithik.Visible = False
imgAkshay.Visible = False
imgAb.Visible = True
imgAjay.Visible = False
imgAmir.Visible = True
imgSmk.Visible = True
imgBobby.Visible = False
imgAkki.Visible = True
imgAnil.Visible = True
cmdList.Caption = "List3"
ElseIf cmdList.Caption = "List3" Then
 Scan = Scan + 2
imgSrk.Visible = False
imgRithik.Visible = False
imgAkshay.Visible = True
imgAb.Visible = False
imgAjay.Visible = True
imgAmir.Visible = True
imgSmk.Visible = False
imgBobby.Visible = True
imgAkki.Visible = False
imgAnil.Visible = True
cmdList.Caption = "List4"
'ElseIf cmdList.Caption = "List4" Then
'Scan = Scan + 1
'imgSrk.Visible = True
'imgRithik.Visible = True
'imgAkshay.Visible = True
'imgAb.Visible = True
'imgAjay.Visible = False
'imgAmir.Visible = False
'imgSmk.Visible = True
'imgBobby.Visible = True
'imgAkki.Visible = False
'imgAnil.Visible = False
'cmdList.Caption = "List5"
ElseIf cmdList.Caption = "List4" Then
 Scan = Scan + 1
imgSrk.Visible = False
imgRithik.Visible = False
imgAkshay.Visible = False
imgAb.Visible = False
imgAjay.Visible = False
imgAmir.Visible = False
imgSmk.Visible = False
imgBobby.Visible = False
imgAkki.Visible = False
imgAnil.Visible = False
cmdList.Visible = False
imgChoice.Visible = False
cmdYes.Visible = False
cmdNo.Visible = False
cmdResult.Visible = True
End If
End Sub
Private Sub cmdResult_Click()
If Scan = 1 Then
imgBobby.Visible = True
ElseIf Scan = 2 Then
imgSmk.Visible = True
ElseIf Scan = 3 Then
imgAnil.Visible = True
ElseIf Scan = 4 Then
imgAjay.Visible = True
ElseIf Scan = 5 Then
imgSrk.Visible = True
ElseIf Scan = 6 Then
imgAkki.Visible = True
ElseIf Scan = 7 Then
imgRithik.Visible = True
ElseIf Scan = 8 Then
imgAkshay.Visible = True
ElseIf Scan = 9 Then
imgAb.Visible = True
ElseIf Scan = 10 Then
imgAmir.Visible = True
ElseIf Scan = 0 Then
imgCom.Visible = True
imgComs.Visible = True
End If
cmdRestart.Visible = True
cmdRestart.Caption = "RESTART"
End Sub

Private Sub Form_Load()
Scan = 0
Start = 0 'make start button false

End Sub

Private Sub imgBg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClose1.Visible = True
imgClose2.Visible = False
imgCredit2.Visible = True
imgCredit1.Visible = False
If Button = 1 Then
 If Y < 400 Then
 frmActors.Left = (frmActors.Left + X) ' - frmActors.Left
 frmActors.Top = (frmActors.Top + Y) '- frmActors.Top+
 End If
End If
End Sub

Private Sub imgClose2_Click()
Load frmMinds
frmMinds.Visible = True
Unload frmActors 'Me
Unload frmActress
End Sub

Private Sub imgClose1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClose1.Visible = False
imgClose2.Visible = True
End Sub

Private Sub imgCredit1_Click()
frmCredit.Visible = True

End Sub

Private Sub imgCredit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCredit1.Visible = True
imgCredit2.Visible = False
End Sub

