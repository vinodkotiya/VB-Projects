VERSION 5.00
Begin VB.Form WELCOME 
   Caption         =   "WELCOME"
   ClientHeight    =   6960
   ClientLeft      =   2415
   ClientTop       =   1290
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   Picture         =   "WELCOME.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   7950
   Begin VB.CommandButton CMD_CONT 
      BackColor       =   &H80000018&
      Caption         =   "CONTINUE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label LBL_WELCOME 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO THE          B.E.C.  ATM                   SERVICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   3255
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "WELCOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_CONT_Click()
Unload Me
LOGINSCREEN.Show

End Sub

