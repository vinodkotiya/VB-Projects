VERSION 5.00
Begin VB.Form LOGOUT 
   Caption         =   "LOGOUT"
   ClientHeight    =   6900
   ClientLeft      =   2550
   ClientTop       =   1290
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   Picture         =   "LOGOUT.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   7575
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   720
      Top             =   3960
   End
   Begin VB.Label LBL_PRJ 
      BackStyle       =   0  'Transparent
      Caption         =   "PROJECT SUBMITTED BY SIDDHARTH TAPARIA     ITI MITTAL               PREETI KHURANA VIJAYETA  KHANDALKAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1935
      Left            =   2160
      TabIndex        =   2
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label LBL_THANKYOU 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "THANK  YOU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label LBL_LOGOUT 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "FOR BANKING WITH B.E.C.  ATM SERVICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   7455
   End
End
Attribute VB_Name = "LOGOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
WELCOME.Show
End Sub
