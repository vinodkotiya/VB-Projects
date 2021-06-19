VERSION 5.00
Begin VB.Form SEE_BALANCE 
   Caption         =   "SEE_BALANCE"
   ClientHeight    =   6915
   ClientLeft      =   2550
   ClientTop       =   1365
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   Picture         =   "SEE_BALANCE.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   7620
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
   Begin VB.Label LBL_BALANCE 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label LBL_SEE_BALANCE 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "SEE_BALANCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_RETURN_Click()
Unload Me
MENUSCREEN.Show
End Sub

Private Sub Form_Load()
LBL_SEE_BALANCE.Caption = "Mr./Ms " & NM
LBL_BALANCE.Caption = "YOUR BALANCE IS Rs. " & BAL
End Sub
