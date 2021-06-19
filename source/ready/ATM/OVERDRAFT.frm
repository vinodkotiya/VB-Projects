VERSION 5.00
Begin VB.Form OVERDRAFT 
   Caption         =   "OVERDRAFT"
   ClientHeight    =   6960
   ClientLeft      =   2550
   ClientTop       =   1365
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   Picture         =   "OVERDRAFT.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   7650
   Begin VB.CommandButton CMD_RETURN 
      BackColor       =   &H80000018&
      Caption         =   "  RETURN TO       MAIN MENU      "
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
   Begin VB.Label LBL_OVERDRAFT 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   $"OVERDRAFT.frx":B5F2
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
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "OVERDRAFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_RETURN_Click()
Unload Me
MENUSCREEN.Show

End Sub
