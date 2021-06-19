VERSION 5.00
Begin VB.Form MENUSCREEN 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MENU"
   ClientHeight    =   6795
   ClientLeft      =   2490
   ClientTop       =   1290
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "MENUSCREEN.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   7620
   Begin VB.CommandButton CMD_LOGOUT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "LOG OUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton CMD_CHANGEPWD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton CMD_MINISTATEMENT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "MINI STATEMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton CMD_BALANCE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "BALANCE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton CMD_DEPOSIT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "DEPOSIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton CMD_WITHDRAW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "WITHDRAW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -240
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label LBL_NM 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label LBL_MENU 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   6375
   End
End
Attribute VB_Name = "MENUSCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_BALANCE_Click()
Unload Me
SEE_BALANCE.Show
End Sub

Private Sub CMD_CHANGEPWD_Click()
Unload Me
CHANGE_PWD.Show
End Sub

Private Sub CMD_DEPOSIT_Click()
Unload Me
DEPOSIT.Show
End Sub

Private Sub CMD_LOGOUT_Click()
Unload Me
LOGOUT.Show
End Sub

Private Sub CMD_WITHDRAW_Click()
Unload Me
WITHDRAWAL.Show
End Sub

Private Sub Form_Load()
LBL_MENU.Caption = "THIS IS THE ACCOUNT OF "
LBL_MENU.Visible = True
LBL_NM.Caption = "Mr.\Ms. " & NM
LBL_NM.Visible = True
End Sub

Private Sub Label1_Click()

End Sub
