VERSION 5.00
Begin VB.Form frmActress 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "MIND SCANNER"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdRestart 
      Caption         =   "START"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "Result"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00FF80FF&
      Caption         =   "List1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgCredit2 
      Height          =   330
      Left            =   8160
      Picture         =   "frmActress.frx":0000
      Top             =   600
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgCredit1 
      Height          =   330
      Left            =   8160
      Picture         =   "frmActress.frx":058D
      Top             =   600
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgComs 
      Height          =   540
      Left            =   1920
      Picture         =   "frmActress.frx":0A6E
      Top             =   3000
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.Image imgCom 
      Height          =   2430
      Left            =   240
      Picture         =   "frmActress.frx":31BF
      Top             =   2280
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgGuide 
      Height          =   1080
      Left            =   120
      Picture         =   "frmActress.frx":4A5B
      Top             =   1200
      Width           =   7605
   End
   Begin VB.Image imgClose2 
      Height          =   300
      Left            =   8740
      Picture         =   "frmActress.frx":52D6
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose1 
      Height          =   300
      Left            =   8740
      Picture         =   "frmActress.frx":5818
      Top             =   90
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   120
      Picture         =   "frmActress.frx":5D5A
      Top             =   470
      Width           =   5160
   End
   Begin VB.Image imgKirty 
      Height          =   1230
      Left            =   5520
      Picture         =   "frmActress.frx":75ED
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Image imgChoice 
      Height          =   300
      Left            =   120
      Picture         =   "frmActress.frx":8FF7
      Top             =   1800
      Visible         =   0   'False
      Width           =   8850
   End
   Begin VB.Image imgPriety 
      Height          =   1230
      Left            =   3000
      Picture         =   "frmActress.frx":C279
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Image imgShilpa 
      Height          =   1230
      Left            =   6000
      Picture         =   "frmActress.frx":DC83
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Image imgMadhuri 
      Height          =   1230
      Left            =   7920
      Picture         =   "frmActress.frx":F68D
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Image imgKareena 
      Height          =   1230
      Left            =   240
      Picture         =   "frmActress.frx":11097
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Image imgRavi 
      Height          =   1230
      Left            =   7920
      Picture         =   "frmActress.frx":12AA1
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Image imgMan 
      Height          =   1230
      Left            =   240
      Picture         =   "frmActress.frx":144AB
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Image imgJuhi 
      Height          =   1230
      Left            =   1440
      Picture         =   "frmActress.frx":15EB5
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Image imgAsh 
      Height          =   1230
      Left            =   3000
      Picture         =   "frmActress.frx":178BF
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Image imgSush 
      Height          =   1230
      Left            =   4560
      Picture         =   "frmActress.frx":192C9
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Image imgBg 
      Height          =   4980
      Left            =   0
      Picture         =   "frmActress.frx":1ACD3
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmActress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Scan As Integer, Start As Integer


Private Sub cmdNo_Click()
If cmdList.Caption = "List1" Then
 Scan = Scan + 0
imgKirty.Visible = False
imgMadhuri.Visible = True
imgAsh.Visible = True
imgKareena.Visible = True
imgJuhi.Visible = False
imgSush.Visible = True
imgPriety.Visible = False
imgRavi.Visible = False
imgShilpa.Visible = True
imgMan.Visible = False
imgAsh.Top = imgAsh.Top + 1230
imgSush.Top = imgSush.Top - 1300
cmdList.Caption = "List2"
ElseIf cmdList.Caption = "List2" Then
 Scan = Scan + 0
imgKirty.Visible = True
imgMadhuri.Visible = False
imgAsh.Visible = False
imgKareena.Visible = True
imgJuhi.Visible = False
imgSush.Visible = True
imgPriety.Visible = True
imgRavi.Visible = False
imgShilpa.Visible = True
imgMan.Visible = True
cmdList.Caption = "List3"
ElseIf cmdList.Caption = "List3" Then
 Scan = Scan + 0
imgKirty.Visible = False
imgMadhuri.Visible = False
imgAsh.Visible = True
imgKareena.Visible = False
imgJuhi.Visible = True
imgSush.Visible = True
imgPriety.Visible = False
imgRavi.Visible = True
imgShilpa.Visible = False
imgMan.Visible = True
imgKareena.Left = 7800
cmdList.Caption = "List4"
'ElseIf cmdList.Caption = "List4" Then
'Scan = Scan + 0
'imgKirty.Visible = True
'imgMadhuri.Visible = True
'imgAsh.Visible = True
'imgKareena.Visible = True
'imgJuhi.Visible = False
'imgSush.Visible = False
'imgPriety.Visible = True
'imgRavi.Visible = True
'imgShilpa.Visible = False
'imgMan.Visible = False
'cmdList.Caption = "List5"

ElseIf cmdList.Caption = "List4" Then
 Scan = Scan + 0
imgKirty.Visible = False
imgMadhuri.Visible = False
imgAsh.Visible = False
imgKareena.Visible = False
imgJuhi.Visible = False
imgSush.Visible = False
imgPriety.Visible = False
imgRavi.Visible = False
imgShilpa.Visible = False
imgMan.Visible = False
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
 imgKirty.Visible = True
 imgMadhuri.Visible = True
 imgAsh.Visible = True
 imgKareena.Visible = True
 imgJuhi.Visible = True
 imgSush.Visible = True
 imgPriety.Visible = False
 imgRavi.Visible = False
 imgShilpa.Visible = False
 imgMan.Visible = False
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
 frmActress.Visible = False
 Load frmMinds
 frmMinds.Visible = True
 Unload frmActors
 Unload frmActress
 End If
End Sub

Private Sub cmdYes_Click()
If cmdList.Caption = "List1" Then
 Scan = Scan + 3           'column 3
imgKirty.Visible = False
imgMadhuri.Visible = True
imgAsh.Visible = True
imgKareena.Visible = True
imgJuhi.Visible = False
imgSush.Visible = True
imgPriety.Visible = False
imgRavi.Visible = False
imgShilpa.Visible = True
imgMan.Visible = False
cmdList.Caption = "List2"
ElseIf cmdList.Caption = "List2" Then
 Scan = Scan + 4
imgKirty.Visible = True
imgMadhuri.Visible = False
imgAsh.Visible = False
imgKareena.Visible = True
imgJuhi.Visible = False
imgSush.Visible = True
imgPriety.Visible = True
imgRavi.Visible = False
imgShilpa.Visible = True
imgMan.Visible = True
cmdList.Caption = "List3"
ElseIf cmdList.Caption = "List3" Then
 Scan = Scan + 2
imgKirty.Visible = False
imgMadhuri.Visible = False
imgAsh.Visible = True
imgKareena.Visible = False
imgJuhi.Visible = True
imgSush.Visible = True
imgPriety.Visible = False
imgRavi.Visible = True
imgShilpa.Visible = False
imgMan.Visible = True
cmdList.Caption = "List4"
'ElseIf cmdList.Caption = "List4" Then
'Scan = Scan + 1
'imgKirty.Visible = True
'imgMadhuri.Visible = True
'imgAsh.Visible = True
'imgKareena.Visible = True
'imgJuhi.Visible = False
'imgSush.Visible = False
'imgPriety.Visible = True
'imgRavi.Visible = True
'imgShilpa.Visible = False
'imgMan.Visible = False
'cmdList.Caption = "List5"
ElseIf cmdList.Caption = "List4" Then
 Scan = Scan + 1
imgKirty.Visible = False
imgMadhuri.Visible = False
imgAsh.Visible = False
imgKareena.Visible = False
imgJuhi.Visible = False
imgSush.Visible = False
imgPriety.Visible = False
imgRavi.Visible = False
imgShilpa.Visible = False
imgMan.Visible = False
cmdList.Visible = False
imgChoice.Visible = False
cmdYes.Visible = False
cmdNo.Visible = False
cmdResult.Visible = True
End If
End Sub
Private Sub cmdResult_Click()
If Scan = 1 Then
imgRavi.Visible = True
ElseIf Scan = 2 Then
imgPriety.Visible = True
ElseIf Scan = 3 Then
imgMan.Visible = True
ElseIf Scan = 4 Then
imgJuhi.Visible = True
ElseIf Scan = 5 Then
imgKirty.Visible = True
ElseIf Scan = 6 Then
imgShilpa.Visible = True
ElseIf Scan = 7 Then
imgMadhuri.Visible = True
ElseIf Scan = 8 Then
imgAsh.Visible = True
ElseIf Scan = 9 Then
imgKareena.Visible = True
ElseIf Scan = 10 Then
imgSush.Visible = True
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
 frmActress.Left = (frmActress.Left + X) ' - frmActress.Left
 frmActress.Top = (frmActress.Top + Y) '- frmActress.Top+
 End If
End If
End Sub

Private Sub imgClose2_Click()
Load frmMinds
frmMinds.Visible = True
Unload frmActress 'Me
Unload frmActors
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
