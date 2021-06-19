VERSION 5.00
Begin VB.Form COLLECT_CASH 
   Caption         =   "COLLECT_CASH"
   ClientHeight    =   6990
   ClientLeft      =   2550
   ClientTop       =   1290
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   Picture         =   "COLLECT_CASH.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   7635
   Begin VB.CommandButton CMD_RETURN 
      BackColor       =   &H80000018&
      Caption         =   "RETURN TO MAIN MENU "
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
   Begin VB.Label LBL_COLLECT_CASH 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "  PLEASE COLLECT YOUR CASH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
   End
End
Attribute VB_Name = "COLLECT_CASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_RETURN_Click()
Unload Me
MENUSCREEN.Show
End Sub
