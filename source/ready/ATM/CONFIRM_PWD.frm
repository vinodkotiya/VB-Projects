VERSION 5.00
Begin VB.Form CONFIRM_PWD 
   Caption         =   "CONFIRM_ PWD"
   ClientHeight    =   6990
   ClientLeft      =   2550
   ClientTop       =   1170
   ClientWidth     =   7635
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "CONFIRM_PWD.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   7635
   Begin VB.CommandButton CMD_RETURN 
      BackColor       =   &H80000018&
      Caption         =   "RETURN TO MAIN MENU"
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
   Begin VB.Label LBL_CONFIRMPASSWORD 
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR PASSWORD HAS BEEN CONFIRMED"
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7455
   End
End
Attribute VB_Name = "CONFIRM_PWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_RETURN_Click()
Unload Me
MENUSCREEN.Show
End Sub
