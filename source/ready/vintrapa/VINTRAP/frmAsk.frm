VERSION 5.00
Begin VB.Form frmAsk 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Registration"
   ClientHeight    =   2355
   ClientLeft      =   3720
   ClientTop       =   2865
   ClientWidth     =   4695
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<Back"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdCredit 
      Caption         =   "CREDIT"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next>>"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtId 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblRok 
      BackStyle       =   0  'Transparent
      Caption         =   "This window will not appear in full version."
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblId 
      BackStyle       =   0  'Transparent
      Caption         =   "FOR FULL VERSION ENTER REGISTRATION CODE  THEN CLICK NEXT"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Image imgYes 
      Height          =   330
      Left            =   1320
      Picture         =   "frmAsk.frx":0000
      Top             =   720
      Width           =   750
   End
   Begin VB.Label lblAsk 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Do you want to quit the current game"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image imgNo 
      Height          =   330
      Left            =   2760
      Picture         =   "frmAsk.frx":0610
      Top             =   720
      Width           =   750
   End
End
Attribute VB_Name = "frmAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub


Private Sub cmdBack_Click()
If cmdBack.Caption = "Quit" Then
 Unload frmCredit
 Unload frmField
Else
 frmField.Visible = True
End If
 Unload Me

End Sub

Private Sub cmdCredit_Click()
Load frmCredit
frmCredit.Visible = True
End Sub

Private Sub cmdNext_Click()
MsgBox "The Registration I.D. You Have Entered Is Incorrect"
End Sub

Private Sub imgNo_Click()
frmField.Visible = True
frmField.timBall.Interval = frmField.timFun.Interval
Unload Me

End Sub

Private Sub imgYes_Click()
Unload frmField
Load frmField
frmField.Visible = True
Unload Me
End Sub

